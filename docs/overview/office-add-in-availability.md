---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448146"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="12827-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="12827-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="12827-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="12827-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="12827-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="12827-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="12827-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="12827-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="12827-108">Excel</span><span class="sxs-lookup"><span data-stu-id="12827-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="12827-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="12827-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="12827-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="12827-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="12827-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="12827-113">Office Online</span></span></td>
    <td> <span data-ttu-id="12827-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-114">- TaskPane</span></span><br><span data-ttu-id="12827-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-115">
        - Content</span></span><br><span data-ttu-id="12827-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="12827-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12827-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-126">
        - BindingEvents</span></span><br><span data-ttu-id="12827-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-127">
        - CompressedFile</span></span><br><span data-ttu-id="12827-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-128">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-129">
        - File</span></span><br><span data-ttu-id="12827-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-130">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-132">
        - Selection</span></span><br><span data-ttu-id="12827-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-133">
        - Settings</span></span><br><span data-ttu-id="12827-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-134">
        - TableBindings</span></span><br><span data-ttu-id="12827-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-135">
        - TableCoercion</span></span><br><span data-ttu-id="12827-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-136">
        - TextBindings</span></span><br><span data-ttu-id="12827-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-138">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-138">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-139">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-139">- TaskPane</span></span><br><span data-ttu-id="12827-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-140">
        - Content</span></span><br><span data-ttu-id="12827-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="12827-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12827-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-151">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-151">
        - BindingEvents</span></span><br><span data-ttu-id="12827-152">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-152">
        - CompressedFile</span></span><br><span data-ttu-id="12827-153">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-153">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-154">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-154">
        - File</span></span><br><span data-ttu-id="12827-155">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-155">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-156">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-156">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-157">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-157">
        - Selection</span></span><br><span data-ttu-id="12827-158">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-158">
        - Settings</span></span><br><span data-ttu-id="12827-159">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-159">
        - TableBindings</span></span><br><span data-ttu-id="12827-160">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-160">
        - TableCoercion</span></span><br><span data-ttu-id="12827-161">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-161">
        - TextBindings</span></span><br><span data-ttu-id="12827-162">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-162">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-163">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-163">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="12827-164">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-164">- TaskPane</span></span><br><span data-ttu-id="12827-165">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-165">
        - Content</span></span><br><span data-ttu-id="12827-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12827-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-176">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-176">- BindingEvents</span></span><br><span data-ttu-id="12827-177">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-177">
        - CompressedFile</span></span><br><span data-ttu-id="12827-178">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-178">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-179">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-179">
        - File</span></span><br><span data-ttu-id="12827-180">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-180">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-181">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-181">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-182">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-182">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-183">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-183">
        - Selection</span></span><br><span data-ttu-id="12827-184">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-184">
        - Settings</span></span><br><span data-ttu-id="12827-185">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-185">
        - TableBindings</span></span><br><span data-ttu-id="12827-186">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-186">
        - TableCoercion</span></span><br><span data-ttu-id="12827-187">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-187">
        - TextBindings</span></span><br><span data-ttu-id="12827-188">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-188">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-189">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-189">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="12827-190">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-190">- TaskPane</span></span><br><span data-ttu-id="12827-191">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-191">
        - Content</span></span></td>
    <td><span data-ttu-id="12827-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="12827-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-194">- BindingEvents</span></span><br><span data-ttu-id="12827-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-195">
        - CompressedFile</span></span><br><span data-ttu-id="12827-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-196">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-197">
        - File</span></span><br><span data-ttu-id="12827-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-198">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-199">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-201">
        - Selection</span></span><br><span data-ttu-id="12827-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-202">
        - Settings</span></span><br><span data-ttu-id="12827-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-203">
        - TableBindings</span></span><br><span data-ttu-id="12827-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-204">
        - TableCoercion</span></span><br><span data-ttu-id="12827-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-205">
        - TextBindings</span></span><br><span data-ttu-id="12827-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-207">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-207">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="12827-208">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-208">
        - TaskPane</span></span><br><span data-ttu-id="12827-209">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-209">
        - Content</span></span></td>
    <td>  <span data-ttu-id="12827-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12827-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="12827-211">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-211">
        - BindingEvents</span></span><br><span data-ttu-id="12827-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-212">
        - CompressedFile</span></span><br><span data-ttu-id="12827-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-213">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-214">
        - File</span></span><br><span data-ttu-id="12827-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-215">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-216">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-218">
        - Selection</span></span><br><span data-ttu-id="12827-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-219">
        - Settings</span></span><br><span data-ttu-id="12827-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-220">
        - TableBindings</span></span><br><span data-ttu-id="12827-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-221">
        - TableCoercion</span></span><br><span data-ttu-id="12827-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-222">
        - TextBindings</span></span><br><span data-ttu-id="12827-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-224">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="12827-224">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="12827-225">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-225">- TaskPane</span></span><br><span data-ttu-id="12827-226">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-226">
        - Content</span></span></td>
    <td><span data-ttu-id="12827-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-236">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-236">- BindingEvents</span></span><br><span data-ttu-id="12827-237">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-237">
        - CompressedFile</span></span><br><span data-ttu-id="12827-238">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-238">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-239">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-239">
        - File</span></span><br><span data-ttu-id="12827-240">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-240">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-241">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-241">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-242">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-242">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-243">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-243">
        - Selection</span></span><br><span data-ttu-id="12827-244">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-244">
        - Settings</span></span><br><span data-ttu-id="12827-245">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-245">
        - TableBindings</span></span><br><span data-ttu-id="12827-246">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-246">
        - TableCoercion</span></span><br><span data-ttu-id="12827-247">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-247">
        - TextBindings</span></span><br><span data-ttu-id="12827-248">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-248">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-249">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-249">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="12827-250">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-250">- TaskPane</span></span><br><span data-ttu-id="12827-251">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-251">
        - Content</span></span><br><span data-ttu-id="12827-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12827-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-262">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-262">- BindingEvents</span></span><br><span data-ttu-id="12827-263">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-263">
        - CompressedFile</span></span><br><span data-ttu-id="12827-264">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-264">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-265">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-265">
        - File</span></span><br><span data-ttu-id="12827-266">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-266">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-267">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-267">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-268">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-268">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-269">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-269">
        - PdfFile</span></span><br><span data-ttu-id="12827-270">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-270">
        - Selection</span></span><br><span data-ttu-id="12827-271">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-271">
        - Settings</span></span><br><span data-ttu-id="12827-272">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-272">
        - TableBindings</span></span><br><span data-ttu-id="12827-273">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-273">
        - TableCoercion</span></span><br><span data-ttu-id="12827-274">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-274">
        - TextBindings</span></span><br><span data-ttu-id="12827-275">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-275">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-276">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-276">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="12827-277">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-277">- TaskPane</span></span><br><span data-ttu-id="12827-278">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-278">
        - Content</span></span><br><span data-ttu-id="12827-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12827-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12827-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12827-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12827-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12827-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12827-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12827-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12827-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12827-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12827-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-289">- BindingEvents</span></span><br><span data-ttu-id="12827-290">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-290">
        - CompressedFile</span></span><br><span data-ttu-id="12827-291">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-291">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-292">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-292">
        - File</span></span><br><span data-ttu-id="12827-293">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-293">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-294">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-294">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-295">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-295">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-296">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-296">
        - PdfFile</span></span><br><span data-ttu-id="12827-297">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-297">
        - Selection</span></span><br><span data-ttu-id="12827-298">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-298">
        - Settings</span></span><br><span data-ttu-id="12827-299">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-299">
        - TableBindings</span></span><br><span data-ttu-id="12827-300">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-300">
        - TableCoercion</span></span><br><span data-ttu-id="12827-301">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-301">
        - TextBindings</span></span><br><span data-ttu-id="12827-302">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-302">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-303">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-303">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="12827-304">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-304">- TaskPane</span></span><br><span data-ttu-id="12827-305">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-305">
        - Content</span></span></td>
    <td><span data-ttu-id="12827-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12827-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="12827-308">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-308">- BindingEvents</span></span><br><span data-ttu-id="12827-309">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-309">
        - CompressedFile</span></span><br><span data-ttu-id="12827-310">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-310">
        - DocumentEvents</span></span><br><span data-ttu-id="12827-311">
        - File</span><span class="sxs-lookup"><span data-stu-id="12827-311">
        - File</span></span><br><span data-ttu-id="12827-312">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-312">
        - ImageCoercion</span></span><br><span data-ttu-id="12827-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-313">
        - MatrixBindings</span></span><br><span data-ttu-id="12827-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="12827-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-315">
        - PdfFile</span></span><br><span data-ttu-id="12827-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-316">
        - Selection</span></span><br><span data-ttu-id="12827-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-317">
        - Settings</span></span><br><span data-ttu-id="12827-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-318">
        - TableBindings</span></span><br><span data-ttu-id="12827-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-319">
        - TableCoercion</span></span><br><span data-ttu-id="12827-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-320">
        - TextBindings</span></span><br><span data-ttu-id="12827-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-321">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12827-322">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="12827-322">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="12827-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="12827-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12827-324">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="12827-324">Platform</span></span></th>
    <th><span data-ttu-id="12827-325">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-325">Extension points</span></span></th>
    <th><span data-ttu-id="12827-326">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-326">API requirement sets</span></span></th>
    <th><span data-ttu-id="12827-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="12827-328">Office Online</span></span></td>
    <td> <span data-ttu-id="12827-329">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-329">- Mail Read</span></span><br><span data-ttu-id="12827-330">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-330">
      - Mail Compose</span></span><br><span data-ttu-id="12827-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12827-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12827-339">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-340">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-340">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-341">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-341">- Mail Read</span></span><br><span data-ttu-id="12827-342">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-342">
      - Mail Compose</span></span><br><span data-ttu-id="12827-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12827-344">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="12827-344">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12827-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12827-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12827-352">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-353">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-353">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-354">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-354">- Mail Read</span></span><br><span data-ttu-id="12827-355">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-355">
      - Mail Compose</span></span><br><span data-ttu-id="12827-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12827-357">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="12827-357">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12827-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12827-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12827-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12827-365">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-366">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-366">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-367">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-367">- Mail Read</span></span><br><span data-ttu-id="12827-368">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-368">
      - Mail Compose</span></span><br><span data-ttu-id="12827-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12827-370">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="12827-370">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12827-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12827-375">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-376">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-376">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-377">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-377">- Mail Read</span></span><br><span data-ttu-id="12827-378">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-378">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="12827-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="12827-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12827-383">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-384">Office 365 pour iOS</span><span class="sxs-lookup"><span data-stu-id="12827-384">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="12827-385">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-385">- Mail Read</span></span><br><span data-ttu-id="12827-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12827-392">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-393">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-393">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-394">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-394">- Mail Read</span></span><br><span data-ttu-id="12827-395">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-395">
      - Mail Compose</span></span><br><span data-ttu-id="12827-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12827-403">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-404">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-404">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-405">- Mail Read</span></span><br><span data-ttu-id="12827-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-406">
      - Mail Compose</span></span><br><span data-ttu-id="12827-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12827-414">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-415">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-415">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-416">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-416">- Mail Read</span></span><br><span data-ttu-id="12827-417">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="12827-417">
      - Mail Compose</span></span><br><span data-ttu-id="12827-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12827-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12827-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12827-425">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-426">Office 365 pour Android</span><span class="sxs-lookup"><span data-stu-id="12827-426">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="12827-427">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="12827-427">- Mail Read</span></span><br><span data-ttu-id="12827-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12827-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12827-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12827-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12827-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12827-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12827-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12827-434">Non disponible</span><span class="sxs-lookup"><span data-stu-id="12827-434">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="12827-435">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="12827-435">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="12827-436">Word</span><span class="sxs-lookup"><span data-stu-id="12827-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12827-437">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="12827-437">Platform</span></span></th>
    <th><span data-ttu-id="12827-438">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-438">Extension points</span></span></th>
    <th><span data-ttu-id="12827-439">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="12827-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="12827-441">Office Online</span></span></td>
    <td> <span data-ttu-id="12827-442">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-442">- TaskPane</span></span><br><span data-ttu-id="12827-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-448">- BindingEvents</span></span><br><span data-ttu-id="12827-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-450">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-451">
         - File</span></span><br><span data-ttu-id="12827-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-453">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-454">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-457">
         - PdfFile</span></span><br><span data-ttu-id="12827-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-458">
         - Selection</span></span><br><span data-ttu-id="12827-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-459">
         - Settings</span></span><br><span data-ttu-id="12827-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-460">
         - TableBindings</span></span><br><span data-ttu-id="12827-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-461">
         - TableCoercion</span></span><br><span data-ttu-id="12827-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-462">
         - TextBindings</span></span><br><span data-ttu-id="12827-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-463">
         - TextCoercion</span></span><br><span data-ttu-id="12827-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-465">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-466">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-466">- TaskPane</span></span><br><span data-ttu-id="12827-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-472">- BindingEvents</span></span><br><span data-ttu-id="12827-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-473">
         - CompressedFile</span></span><br><span data-ttu-id="12827-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-475">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-476">
         - File</span></span><br><span data-ttu-id="12827-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-478">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-479">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-482">
         - PdfFile</span></span><br><span data-ttu-id="12827-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-483">
         - Selection</span></span><br><span data-ttu-id="12827-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-484">
         - Settings</span></span><br><span data-ttu-id="12827-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-485">
         - TableBindings</span></span><br><span data-ttu-id="12827-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-486">
         - TableCoercion</span></span><br><span data-ttu-id="12827-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-487">
         - TextBindings</span></span><br><span data-ttu-id="12827-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-488">
         - TextCoercion</span></span><br><span data-ttu-id="12827-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-490">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-491">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-491">- TaskPane</span></span><br><span data-ttu-id="12827-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-497">- BindingEvents</span></span><br><span data-ttu-id="12827-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-498">
         - CompressedFile</span></span><br><span data-ttu-id="12827-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-500">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-501">
         - File</span></span><br><span data-ttu-id="12827-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-503">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-504">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-507">
         - PdfFile</span></span><br><span data-ttu-id="12827-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-508">
         - Selection</span></span><br><span data-ttu-id="12827-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-509">
         - Settings</span></span><br><span data-ttu-id="12827-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-510">
         - TableBindings</span></span><br><span data-ttu-id="12827-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-511">
         - TableCoercion</span></span><br><span data-ttu-id="12827-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-512">
         - TextBindings</span></span><br><span data-ttu-id="12827-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-513">
         - TextCoercion</span></span><br><span data-ttu-id="12827-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-515">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-516">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="12827-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-519">- BindingEvents</span></span><br><span data-ttu-id="12827-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-520">
         - CompressedFile</span></span><br><span data-ttu-id="12827-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-522">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-523">
         - File</span></span><br><span data-ttu-id="12827-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-525">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-526">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-529">
         - PdfFile</span></span><br><span data-ttu-id="12827-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-530">
         - Selection</span></span><br><span data-ttu-id="12827-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-531">
         - Settings</span></span><br><span data-ttu-id="12827-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-532">
         - TableBindings</span></span><br><span data-ttu-id="12827-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-533">
         - TableCoercion</span></span><br><span data-ttu-id="12827-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-534">
         - TextBindings</span></span><br><span data-ttu-id="12827-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-535">
         - TextCoercion</span></span><br><span data-ttu-id="12827-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-537">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-538">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12827-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12827-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-540">- BindingEvents</span></span><br><span data-ttu-id="12827-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-541">
         - CompressedFile</span></span><br><span data-ttu-id="12827-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-543">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-544">
         - File</span></span><br><span data-ttu-id="12827-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-546">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-547">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-550">
         - PdfFile</span></span><br><span data-ttu-id="12827-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-551">
         - Selection</span></span><br><span data-ttu-id="12827-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-552">
         - Settings</span></span><br><span data-ttu-id="12827-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-553">
         - TableBindings</span></span><br><span data-ttu-id="12827-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-554">
         - TableCoercion</span></span><br><span data-ttu-id="12827-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-555">
         - TextBindings</span></span><br><span data-ttu-id="12827-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-556">
         - TextCoercion</span></span><br><span data-ttu-id="12827-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-558">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="12827-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="12827-559">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12827-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12827-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-564">- BindingEvents</span></span><br><span data-ttu-id="12827-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-565">
         - CompressedFile</span></span><br><span data-ttu-id="12827-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-567">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-568">
         - File</span></span><br><span data-ttu-id="12827-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-570">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-571">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-574">
         - PdfFile</span></span><br><span data-ttu-id="12827-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-575">
         - Selection</span></span><br><span data-ttu-id="12827-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-576">
         - Settings</span></span><br><span data-ttu-id="12827-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-577">
         - TableBindings</span></span><br><span data-ttu-id="12827-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-578">
         - TableCoercion</span></span><br><span data-ttu-id="12827-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-579">
         - TextBindings</span></span><br><span data-ttu-id="12827-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-580">
         - TextCoercion</span></span><br><span data-ttu-id="12827-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-582">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-583">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-583">- TaskPane</span></span><br><span data-ttu-id="12827-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12827-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12827-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-589">- BindingEvents</span></span><br><span data-ttu-id="12827-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-590">
         - CompressedFile</span></span><br><span data-ttu-id="12827-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-592">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-593">
         - File</span></span><br><span data-ttu-id="12827-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-595">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-596">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-599">
         - PdfFile</span></span><br><span data-ttu-id="12827-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-600">
         - Selection</span></span><br><span data-ttu-id="12827-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-601">
         - Settings</span></span><br><span data-ttu-id="12827-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-602">
         - TableBindings</span></span><br><span data-ttu-id="12827-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-603">
         - TableCoercion</span></span><br><span data-ttu-id="12827-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-604">
         - TextBindings</span></span><br><span data-ttu-id="12827-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-605">
         - TextCoercion</span></span><br><span data-ttu-id="12827-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-607">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-608">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-608">- TaskPane</span></span><br><span data-ttu-id="12827-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12827-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12827-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12827-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12827-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12827-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12827-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-614">- BindingEvents</span></span><br><span data-ttu-id="12827-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-615">
         - CompressedFile</span></span><br><span data-ttu-id="12827-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-617">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-618">
         - File</span></span><br><span data-ttu-id="12827-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-620">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-621">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-624">
         - PdfFile</span></span><br><span data-ttu-id="12827-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-625">
         - Selection</span></span><br><span data-ttu-id="12827-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-626">
         - Settings</span></span><br><span data-ttu-id="12827-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-627">
         - TableBindings</span></span><br><span data-ttu-id="12827-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-628">
         - TableCoercion</span></span><br><span data-ttu-id="12827-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-629">
         - TextBindings</span></span><br><span data-ttu-id="12827-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-630">
         - TextCoercion</span></span><br><span data-ttu-id="12827-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-632">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-633">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12827-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12827-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="12827-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12827-636">- BindingEvents</span></span><br><span data-ttu-id="12827-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-637">
         - CompressedFile</span></span><br><span data-ttu-id="12827-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12827-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="12827-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-639">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-640">
         - File</span></span><br><span data-ttu-id="12827-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-642">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12827-643">
         - MatrixBindings</span></span><br><span data-ttu-id="12827-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="12827-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12827-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-646">
         - PdfFile</span></span><br><span data-ttu-id="12827-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-647">
         - Selection</span></span><br><span data-ttu-id="12827-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-648">
         - Settings</span></span><br><span data-ttu-id="12827-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12827-649">
         - TableBindings</span></span><br><span data-ttu-id="12827-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-650">
         - TableCoercion</span></span><br><span data-ttu-id="12827-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12827-651">
         - TextBindings</span></span><br><span data-ttu-id="12827-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-652">
         - TextCoercion</span></span><br><span data-ttu-id="12827-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12827-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="12827-654">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="12827-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="12827-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="12827-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12827-656">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="12827-656">Platform</span></span></th>
    <th><span data-ttu-id="12827-657">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-657">Extension points</span></span></th>
    <th><span data-ttu-id="12827-658">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="12827-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="12827-660">Office Online</span></span></td>
    <td> <span data-ttu-id="12827-661">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-661">- Content</span></span><br><span data-ttu-id="12827-662">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-662">
         - TaskPane</span></span><br><span data-ttu-id="12827-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-665">- ActiveView</span></span><br><span data-ttu-id="12827-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-666">
         - CompressedFile</span></span><br><span data-ttu-id="12827-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-667">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-668">
         - File</span></span><br><span data-ttu-id="12827-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-669">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-670">
         - PdfFile</span></span><br><span data-ttu-id="12827-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-671">
         - Selection</span></span><br><span data-ttu-id="12827-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-672">
         - Settings</span></span><br><span data-ttu-id="12827-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-674">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-675">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-675">- Content</span></span><br><span data-ttu-id="12827-676">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-676">
         - TaskPane</span></span><br><span data-ttu-id="12827-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-679">- ActiveView</span></span><br><span data-ttu-id="12827-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-680">
         - CompressedFile</span></span><br><span data-ttu-id="12827-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-681">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-682">
         - File</span></span><br><span data-ttu-id="12827-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-683">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-684">
         - PdfFile</span></span><br><span data-ttu-id="12827-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-685">
         - Selection</span></span><br><span data-ttu-id="12827-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-686">
         - Settings</span></span><br><span data-ttu-id="12827-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-688">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-689">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-689">- Content</span></span><br><span data-ttu-id="12827-690">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-690">
         - TaskPane</span></span><br><span data-ttu-id="12827-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-693">- ActiveView</span></span><br><span data-ttu-id="12827-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-694">
         - CompressedFile</span></span><br><span data-ttu-id="12827-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-695">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-696">
         - File</span></span><br><span data-ttu-id="12827-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-697">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-698">
         - PdfFile</span></span><br><span data-ttu-id="12827-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-699">
         - Selection</span></span><br><span data-ttu-id="12827-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-700">
         - Settings</span></span><br><span data-ttu-id="12827-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-702">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-703">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-703">- Content</span></span><br><span data-ttu-id="12827-704">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12827-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12827-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-706">- ActiveView</span></span><br><span data-ttu-id="12827-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-707">
         - CompressedFile</span></span><br><span data-ttu-id="12827-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-708">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-709">
         - File</span></span><br><span data-ttu-id="12827-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-710">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-711">
         - PdfFile</span></span><br><span data-ttu-id="12827-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-712">
         - Selection</span></span><br><span data-ttu-id="12827-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-713">
         - Settings</span></span><br><span data-ttu-id="12827-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-715">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-716">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-716">- Content</span></span><br><span data-ttu-id="12827-717">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="12827-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12827-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12827-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-719">- ActiveView</span></span><br><span data-ttu-id="12827-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-720">
         - CompressedFile</span></span><br><span data-ttu-id="12827-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-721">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-722">
         - File</span></span><br><span data-ttu-id="12827-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-723">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-724">
         - PdfFile</span></span><br><span data-ttu-id="12827-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-725">
         - Selection</span></span><br><span data-ttu-id="12827-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-726">
         - Settings</span></span><br><span data-ttu-id="12827-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-728">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="12827-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="12827-729">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-729">- Content</span></span><br><span data-ttu-id="12827-730">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="12827-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-732">- ActiveView</span></span><br><span data-ttu-id="12827-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-733">
         - CompressedFile</span></span><br><span data-ttu-id="12827-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-734">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-735">
         - File</span></span><br><span data-ttu-id="12827-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-736">
         - PdfFile</span></span><br><span data-ttu-id="12827-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-737">
         - Selection</span></span><br><span data-ttu-id="12827-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-738">
         - Settings</span></span><br><span data-ttu-id="12827-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-739">
         - TextCoercion</span></span><br><span data-ttu-id="12827-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-741">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-742">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-742">- Content</span></span><br><span data-ttu-id="12827-743">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-743">
         - TaskPane</span></span><br><span data-ttu-id="12827-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-746">- ActiveView</span></span><br><span data-ttu-id="12827-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-747">
         - CompressedFile</span></span><br><span data-ttu-id="12827-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-748">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-749">
         - File</span></span><br><span data-ttu-id="12827-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-750">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-751">
         - PdfFile</span></span><br><span data-ttu-id="12827-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-752">
         - Selection</span></span><br><span data-ttu-id="12827-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-753">
         - Settings</span></span><br><span data-ttu-id="12827-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-755">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-756">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-756">- Content</span></span><br><span data-ttu-id="12827-757">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-757">
         - TaskPane</span></span><br><span data-ttu-id="12827-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-760">- ActiveView</span></span><br><span data-ttu-id="12827-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-761">
         - CompressedFile</span></span><br><span data-ttu-id="12827-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-762">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-763">
         - File</span></span><br><span data-ttu-id="12827-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-764">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-765">
         - PdfFile</span></span><br><span data-ttu-id="12827-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-766">
         - Selection</span></span><br><span data-ttu-id="12827-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-767">
         - Settings</span></span><br><span data-ttu-id="12827-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-769">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="12827-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="12827-770">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-770">- Content</span></span><br><span data-ttu-id="12827-771">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-771">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12827-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12827-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12827-773">- ActiveView</span></span><br><span data-ttu-id="12827-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12827-774">
         - CompressedFile</span></span><br><span data-ttu-id="12827-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-775">
         - DocumentEvents</span></span><br><span data-ttu-id="12827-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="12827-776">
         - File</span></span><br><span data-ttu-id="12827-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-777">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12827-778">
         - PdfFile</span></span><br><span data-ttu-id="12827-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12827-779">
         - Selection</span></span><br><span data-ttu-id="12827-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-780">
         - Settings</span></span><br><span data-ttu-id="12827-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12827-782">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="12827-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="12827-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="12827-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12827-784">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="12827-784">Platform</span></span></th>
    <th><span data-ttu-id="12827-785">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-785">Extension points</span></span></th>
    <th><span data-ttu-id="12827-786">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="12827-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="12827-788">Office Online</span></span></td>
    <td> <span data-ttu-id="12827-789">- Contenu</span><span class="sxs-lookup"><span data-stu-id="12827-789">- Content</span></span><br><span data-ttu-id="12827-790">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-790">
         - TaskPane</span></span><br><span data-ttu-id="12827-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="12827-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12827-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="12827-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12827-794">- DocumentEvents</span></span><br><span data-ttu-id="12827-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="12827-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-796">
         - ImageCoercion</span></span><br><span data-ttu-id="12827-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12827-797">
         - Settings</span></span><br><span data-ttu-id="12827-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="12827-799">Projet</span><span class="sxs-lookup"><span data-stu-id="12827-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12827-800">Plateforme</span><span class="sxs-lookup"><span data-stu-id="12827-800">Platform</span></span></th>
    <th><span data-ttu-id="12827-801">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="12827-801">Extension points</span></span></th>
    <th><span data-ttu-id="12827-802">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="12827-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="12827-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="12827-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-804">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-805">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="12827-807">- Selection</span></span><br><span data-ttu-id="12827-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-809">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-810">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="12827-812">- Selection</span></span><br><span data-ttu-id="12827-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12827-814">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="12827-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="12827-815">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="12827-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12827-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12827-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12827-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="12827-817">- Selection</span></span><br><span data-ttu-id="12827-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12827-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="12827-819">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="12827-819">See also</span></span>

- [<span data-ttu-id="12827-820">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="12827-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="12827-821">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="12827-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="12827-822">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="12827-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="12827-823">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="12827-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="12827-824">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="12827-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="12827-825">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="12827-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="12827-826">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12827-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="12827-827">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12827-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)