---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 19f2fa7f744345823c2700b04524ec20705035a8
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952368"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="17694-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="17694-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="17694-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="17694-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="17694-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="17694-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="17694-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="17694-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="17694-108">Excel</span><span class="sxs-lookup"><span data-stu-id="17694-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="17694-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="17694-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="17694-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="17694-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="17694-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-113">Office Online</span></span></td>
    <td> <span data-ttu-id="17694-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-114">- TaskPane</span></span><br><span data-ttu-id="17694-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-115">
        - Content</span></span><br><span data-ttu-id="17694-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-116">
        -Custom Functions</span></span><br><span data-ttu-id="17694-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="17694-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="17694-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17694-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17694-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-128">
        - BindingEvents</span></span><br><span data-ttu-id="17694-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-129">
        - CompressedFile</span></span><br><span data-ttu-id="17694-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-130">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-131">
        - File</span></span><br><span data-ttu-id="17694-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-132">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-134">
        - Selection</span></span><br><span data-ttu-id="17694-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-135">
        - Settings</span></span><br><span data-ttu-id="17694-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-136">
        - TableBindings</span></span><br><span data-ttu-id="17694-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-137">
        - TableCoercion</span></span><br><span data-ttu-id="17694-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-138">
        - TextBindings</span></span><br><span data-ttu-id="17694-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-140">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="17694-140">Office apps on Windows</span></span><br><span data-ttu-id="17694-141">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-142">- TaskPane</span></span><br><span data-ttu-id="17694-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-143">
        - Content</span></span><br><span data-ttu-id="17694-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-144">
        -Custom Functions</span></span><br><span data-ttu-id="17694-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="17694-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="17694-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17694-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17694-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-156">
        - BindingEvents</span></span><br><span data-ttu-id="17694-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-157">
        - CompressedFile</span></span><br><span data-ttu-id="17694-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-158">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-159">
        - File</span></span><br><span data-ttu-id="17694-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-160">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-162">
        - Selection</span></span><br><span data-ttu-id="17694-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-163">
        - Settings</span></span><br><span data-ttu-id="17694-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-164">
        - TableBindings</span></span><br><span data-ttu-id="17694-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-165">
        - TableCoercion</span></span><br><span data-ttu-id="17694-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-166">
        - TextBindings</span></span><br><span data-ttu-id="17694-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-168">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-168">Office 2019 for Windows</span></span><br><span data-ttu-id="17694-169">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17694-170">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-170">- TaskPane</span></span><br><span data-ttu-id="17694-171">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-171">
        - Content</span></span><br><span data-ttu-id="17694-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17694-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-182">- BindingEvents</span></span><br><span data-ttu-id="17694-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-183">
        - CompressedFile</span></span><br><span data-ttu-id="17694-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-184">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-185">
        - File</span></span><br><span data-ttu-id="17694-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-186">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-187">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-189">
        - Selection</span></span><br><span data-ttu-id="17694-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-190">
        - Settings</span></span><br><span data-ttu-id="17694-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-191">
        - TableBindings</span></span><br><span data-ttu-id="17694-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-192">
        - TableCoercion</span></span><br><span data-ttu-id="17694-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-193">
        - TextBindings</span></span><br><span data-ttu-id="17694-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-195">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-195">Office 2016 for Windows</span></span><br><span data-ttu-id="17694-196">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17694-197">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-197">- TaskPane</span></span><br><span data-ttu-id="17694-198">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-198">
        - Content</span></span></td>
    <td><span data-ttu-id="17694-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="17694-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-201">- BindingEvents</span></span><br><span data-ttu-id="17694-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-202">
        - CompressedFile</span></span><br><span data-ttu-id="17694-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-203">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-204">
        - File</span></span><br><span data-ttu-id="17694-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-205">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-206">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-208">
        - Selection</span></span><br><span data-ttu-id="17694-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-209">
        - Settings</span></span><br><span data-ttu-id="17694-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-210">
        - TableBindings</span></span><br><span data-ttu-id="17694-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-211">
        - TableCoercion</span></span><br><span data-ttu-id="17694-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-212">
        - TextBindings</span></span><br><span data-ttu-id="17694-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-214">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-214">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="17694-215">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17694-216">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-216">
        - TaskPane</span></span><br><span data-ttu-id="17694-217">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="17694-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17694-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="17694-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-219">
        - BindingEvents</span></span><br><span data-ttu-id="17694-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-220">
        - CompressedFile</span></span><br><span data-ttu-id="17694-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-221">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-222">
        - File</span></span><br><span data-ttu-id="17694-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-223">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-224">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-226">
        - Selection</span></span><br><span data-ttu-id="17694-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-227">
        - Settings</span></span><br><span data-ttu-id="17694-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-228">
        - TableBindings</span></span><br><span data-ttu-id="17694-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-229">
        - TableCoercion</span></span><br><span data-ttu-id="17694-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-230">
        - TextBindings</span></span><br><span data-ttu-id="17694-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-232">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="17694-232">Office for iPad</span></span><br><span data-ttu-id="17694-233">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17694-234">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-234">- TaskPane</span></span><br><span data-ttu-id="17694-235">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-235">
        - Content</span></span><br><span data-ttu-id="17694-236">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-236">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="17694-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17694-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17694-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-247">- BindingEvents</span></span><br><span data-ttu-id="17694-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-248">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-249">
        - File</span></span><br><span data-ttu-id="17694-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-250">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-251">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-253">
        - Selection</span></span><br><span data-ttu-id="17694-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-254">
        - Settings</span></span><br><span data-ttu-id="17694-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-255">
        - TableBindings</span></span><br><span data-ttu-id="17694-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-256">
        - TableCoercion</span></span><br><span data-ttu-id="17694-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-257">
        - TextBindings</span></span><br><span data-ttu-id="17694-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-259">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-259">Office for Mac</span></span><br><span data-ttu-id="17694-260">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17694-261">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-261">- TaskPane</span></span><br><span data-ttu-id="17694-262">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-262">
        - Content</span></span><br><span data-ttu-id="17694-263">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-263">
        -Custom Functions</span></span><br><span data-ttu-id="17694-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17694-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17694-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17694-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-275">- BindingEvents</span></span><br><span data-ttu-id="17694-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-276">
        - CompressedFile</span></span><br><span data-ttu-id="17694-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-277">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-278">
        - File</span></span><br><span data-ttu-id="17694-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-279">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-280">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-282">
        - PdfFile</span></span><br><span data-ttu-id="17694-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-283">
        - Selection</span></span><br><span data-ttu-id="17694-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-284">
        - Settings</span></span><br><span data-ttu-id="17694-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-285">
        - TableBindings</span></span><br><span data-ttu-id="17694-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-286">
        - TableCoercion</span></span><br><span data-ttu-id="17694-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-287">
        - TextBindings</span></span><br><span data-ttu-id="17694-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-289">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-289">Office 2019 for Mac</span></span><br><span data-ttu-id="17694-290">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17694-291">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-291">- TaskPane</span></span><br><span data-ttu-id="17694-292">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-292">
        - Content</span></span><br><span data-ttu-id="17694-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17694-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17694-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17694-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17694-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17694-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17694-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17694-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17694-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17694-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="17694-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-303">- BindingEvents</span></span><br><span data-ttu-id="17694-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-304">
        - CompressedFile</span></span><br><span data-ttu-id="17694-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-305">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-306">
        - File</span></span><br><span data-ttu-id="17694-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-307">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-308">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-310">
        - PdfFile</span></span><br><span data-ttu-id="17694-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-311">
        - Selection</span></span><br><span data-ttu-id="17694-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-312">
        - Settings</span></span><br><span data-ttu-id="17694-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-313">
        - TableBindings</span></span><br><span data-ttu-id="17694-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-314">
        - TableCoercion</span></span><br><span data-ttu-id="17694-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-315">
        - TextBindings</span></span><br><span data-ttu-id="17694-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-317">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-317">Office 2016 for Mac</span></span><br><span data-ttu-id="17694-318">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17694-319">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-319">- TaskPane</span></span><br><span data-ttu-id="17694-320">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-320">
        - Content</span></span></td>
    <td><span data-ttu-id="17694-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17694-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="17694-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-323">- BindingEvents</span></span><br><span data-ttu-id="17694-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-324">
        - CompressedFile</span></span><br><span data-ttu-id="17694-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-325">
        - DocumentEvents</span></span><br><span data-ttu-id="17694-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="17694-326">
        - File</span></span><br><span data-ttu-id="17694-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-327">
        - ImageCoercion</span></span><br><span data-ttu-id="17694-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-328">
        - MatrixBindings</span></span><br><span data-ttu-id="17694-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="17694-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-330">
        - PdfFile</span></span><br><span data-ttu-id="17694-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-331">
        - Selection</span></span><br><span data-ttu-id="17694-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-332">
        - Settings</span></span><br><span data-ttu-id="17694-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-333">
        - TableBindings</span></span><br><span data-ttu-id="17694-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-334">
        - TableCoercion</span></span><br><span data-ttu-id="17694-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-335">
        - TextBindings</span></span><br><span data-ttu-id="17694-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="17694-337">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="17694-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="17694-338">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="17694-339">Plateforme</span><span class="sxs-lookup"><span data-stu-id="17694-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="17694-340">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="17694-341">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="17694-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-343">Office Online</span></span></td>
    <td><span data-ttu-id="17694-344">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-344">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="17694-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-346">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="17694-346">Office apps on Windows</span></span><br><span data-ttu-id="17694-347">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17694-348">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-348">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="17694-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-350">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="17694-350">Office for iPad</span></span><br><span data-ttu-id="17694-351">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17694-352">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-352">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="17694-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-354">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-354">Office for Mac</span></span><br><span data-ttu-id="17694-355">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17694-356">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="17694-356">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="17694-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="17694-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="17694-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17694-359">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="17694-359">Platform</span></span></th>
    <th><span data-ttu-id="17694-360">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-360">Extension points</span></span></th>
    <th><span data-ttu-id="17694-361">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="17694-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-363">Office Online</span></span></td>
    <td> <span data-ttu-id="17694-364">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-364">- Mail Read</span></span><br><span data-ttu-id="17694-365">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-365">
      - Mail Compose</span></span><br><span data-ttu-id="17694-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17694-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17694-374">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-375">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="17694-375">Office apps on Windows</span></span><br><span data-ttu-id="17694-376">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-377">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-377">- Mail Read</span></span><br><span data-ttu-id="17694-378">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-378">
      - Mail Compose</span></span><br><span data-ttu-id="17694-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17694-380">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="17694-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17694-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17694-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17694-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-389">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-389">Office 2019 for Windows</span></span><br><span data-ttu-id="17694-390">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-391">- Mail Read</span></span><br><span data-ttu-id="17694-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-392">
      - Mail Compose</span></span><br><span data-ttu-id="17694-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17694-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="17694-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17694-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17694-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17694-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17694-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-403">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-403">Office 2016 for Windows</span></span><br><span data-ttu-id="17694-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-405">- Mail Read</span></span><br><span data-ttu-id="17694-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-406">
      - Mail Compose</span></span><br><span data-ttu-id="17694-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17694-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="17694-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17694-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="17694-413">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-414">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-414">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="17694-415">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-416">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-416">- Mail Read</span></span><br><span data-ttu-id="17694-417">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="17694-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="17694-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="17694-422">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-423">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="17694-423">Office for iOS</span></span><br><span data-ttu-id="17694-424">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-425">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-425">- Mail Read</span></span><br><span data-ttu-id="17694-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="17694-432">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-433">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-433">Office for Mac</span></span><br><span data-ttu-id="17694-434">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-435">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-435">- Mail Read</span></span><br><span data-ttu-id="17694-436">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-436">
      - Mail Compose</span></span><br><span data-ttu-id="17694-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17694-444">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-445">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-445">Office 2019 for Mac</span></span><br><span data-ttu-id="17694-446">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-447">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-447">- Mail Read</span></span><br><span data-ttu-id="17694-448">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-448">
      - Mail Compose</span></span><br><span data-ttu-id="17694-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17694-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-457">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-457">Office 2016 for Mac</span></span><br><span data-ttu-id="17694-458">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-459">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-459">- Mail Read</span></span><br><span data-ttu-id="17694-460">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="17694-460">
      - Mail Compose</span></span><br><span data-ttu-id="17694-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17694-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17694-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17694-468">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-469">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="17694-469">Office for Android</span></span><br><span data-ttu-id="17694-470">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-470">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-471">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="17694-471">- Mail Read</span></span><br><span data-ttu-id="17694-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17694-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17694-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17694-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17694-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17694-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17694-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="17694-478">Non disponible</span><span class="sxs-lookup"><span data-stu-id="17694-478">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="17694-479">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="17694-479">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="17694-480">Word</span><span class="sxs-lookup"><span data-stu-id="17694-480">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17694-481">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="17694-481">Platform</span></span></th>
    <th><span data-ttu-id="17694-482">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-482">Extension points</span></span></th>
    <th><span data-ttu-id="17694-483">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-483">API requirement sets</span></span></th>
    <th><span data-ttu-id="17694-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-485">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-485">Office Online</span></span></td>
    <td> <span data-ttu-id="17694-486">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-486">- TaskPane</span></span><br><span data-ttu-id="17694-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-492">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-492">- BindingEvents</span></span><br><span data-ttu-id="17694-493">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-493">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-494">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-494">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-495">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-495">
         - File</span></span><br><span data-ttu-id="17694-496">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-496">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-497">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-497">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-498">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-498">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-499">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-499">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-500">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-500">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-501">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-501">
         - PdfFile</span></span><br><span data-ttu-id="17694-502">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-502">
         - Selection</span></span><br><span data-ttu-id="17694-503">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-503">
         - Settings</span></span><br><span data-ttu-id="17694-504">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-504">
         - TableBindings</span></span><br><span data-ttu-id="17694-505">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-505">
         - TableCoercion</span></span><br><span data-ttu-id="17694-506">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-506">
         - TextBindings</span></span><br><span data-ttu-id="17694-507">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-507">
         - TextCoercion</span></span><br><span data-ttu-id="17694-508">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-508">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-509">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="17694-509">Office apps on Windows</span></span><br><span data-ttu-id="17694-510">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-510">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-511">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-511">- TaskPane</span></span><br><span data-ttu-id="17694-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-517">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-517">- BindingEvents</span></span><br><span data-ttu-id="17694-518">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-518">
         - CompressedFile</span></span><br><span data-ttu-id="17694-519">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-519">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-520">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-520">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-521">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-521">
         - File</span></span><br><span data-ttu-id="17694-522">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-522">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-523">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-523">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-524">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-524">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-525">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-525">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-526">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-526">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-527">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-527">
         - PdfFile</span></span><br><span data-ttu-id="17694-528">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-528">
         - Selection</span></span><br><span data-ttu-id="17694-529">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-529">
         - Settings</span></span><br><span data-ttu-id="17694-530">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-530">
         - TableBindings</span></span><br><span data-ttu-id="17694-531">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-531">
         - TableCoercion</span></span><br><span data-ttu-id="17694-532">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-532">
         - TextBindings</span></span><br><span data-ttu-id="17694-533">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-533">
         - TextCoercion</span></span><br><span data-ttu-id="17694-534">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-534">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-535">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-535">Office 2019 for Windows</span></span><br><span data-ttu-id="17694-536">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-536">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-537">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="17694-537">- TaskPane</span></span><br><span data-ttu-id="17694-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-543">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-543">- BindingEvents</span></span><br><span data-ttu-id="17694-544">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-544">
         - CompressedFile</span></span><br><span data-ttu-id="17694-545">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-545">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-546">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-546">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-547">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-547">
         - File</span></span><br><span data-ttu-id="17694-548">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-548">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-549">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-549">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-550">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-550">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-551">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-551">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-552">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-552">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-553">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-553">
         - PdfFile</span></span><br><span data-ttu-id="17694-554">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-554">
         - Selection</span></span><br><span data-ttu-id="17694-555">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-555">
         - Settings</span></span><br><span data-ttu-id="17694-556">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-556">
         - TableBindings</span></span><br><span data-ttu-id="17694-557">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-557">
         - TableCoercion</span></span><br><span data-ttu-id="17694-558">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-558">
         - TextBindings</span></span><br><span data-ttu-id="17694-559">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-559">
         - TextCoercion</span></span><br><span data-ttu-id="17694-560">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-560">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-561">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-561">Office 2016 for Windows</span></span><br><span data-ttu-id="17694-562">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-562">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-563">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-563">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="17694-566">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-566">- BindingEvents</span></span><br><span data-ttu-id="17694-567">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-567">
         - CompressedFile</span></span><br><span data-ttu-id="17694-568">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-568">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-569">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-570">
         - File</span></span><br><span data-ttu-id="17694-571">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-571">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-572">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-572">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-573">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-573">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-574">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-574">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-575">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-575">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-576">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-576">
         - PdfFile</span></span><br><span data-ttu-id="17694-577">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-577">
         - Selection</span></span><br><span data-ttu-id="17694-578">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-578">
         - Settings</span></span><br><span data-ttu-id="17694-579">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-579">
         - TableBindings</span></span><br><span data-ttu-id="17694-580">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-580">
         - TableCoercion</span></span><br><span data-ttu-id="17694-581">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-581">
         - TextBindings</span></span><br><span data-ttu-id="17694-582">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-582">
         - TextCoercion</span></span><br><span data-ttu-id="17694-583">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-583">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-584">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-584">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="17694-585">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-585">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-586">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-586">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17694-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="17694-588">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-588">- BindingEvents</span></span><br><span data-ttu-id="17694-589">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-589">
         - CompressedFile</span></span><br><span data-ttu-id="17694-590">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-590">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-591">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-592">
         - File</span></span><br><span data-ttu-id="17694-593">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-593">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-594">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-595">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-595">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-596">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-596">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-597">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-597">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-598">
         - PdfFile</span></span><br><span data-ttu-id="17694-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-599">
         - Selection</span></span><br><span data-ttu-id="17694-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-600">
         - Settings</span></span><br><span data-ttu-id="17694-601">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-601">
         - TableBindings</span></span><br><span data-ttu-id="17694-602">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-602">
         - TableCoercion</span></span><br><span data-ttu-id="17694-603">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-603">
         - TextBindings</span></span><br><span data-ttu-id="17694-604">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-604">
         - TextCoercion</span></span><br><span data-ttu-id="17694-605">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-605">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-606">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="17694-606">Office for iPad</span></span><br><span data-ttu-id="17694-607">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-607">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-608">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-608">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="17694-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="17694-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-613">- BindingEvents</span></span><br><span data-ttu-id="17694-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-614">
         - CompressedFile</span></span><br><span data-ttu-id="17694-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-616">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-617">
         - File</span></span><br><span data-ttu-id="17694-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-619">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-619">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-620">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-623">
         - PdfFile</span></span><br><span data-ttu-id="17694-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-624">
         - Selection</span></span><br><span data-ttu-id="17694-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-625">
         - Settings</span></span><br><span data-ttu-id="17694-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-626">
         - TableBindings</span></span><br><span data-ttu-id="17694-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-627">
         - TableCoercion</span></span><br><span data-ttu-id="17694-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-628">
         - TextBindings</span></span><br><span data-ttu-id="17694-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-629">
         - TextCoercion</span></span><br><span data-ttu-id="17694-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-631">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-631">Office for Mac</span></span><br><span data-ttu-id="17694-632">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-632">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-633">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-633">- TaskPane</span></span><br><span data-ttu-id="17694-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="17694-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="17694-639">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-639">- BindingEvents</span></span><br><span data-ttu-id="17694-640">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-640">
         - CompressedFile</span></span><br><span data-ttu-id="17694-641">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-641">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-642">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-642">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-643">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-643">
         - File</span></span><br><span data-ttu-id="17694-644">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-644">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-645">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-645">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-646">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-646">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-647">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-647">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-648">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-648">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-649">
         - PdfFile</span></span><br><span data-ttu-id="17694-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-650">
         - Selection</span></span><br><span data-ttu-id="17694-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-651">
         - Settings</span></span><br><span data-ttu-id="17694-652">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-652">
         - TableBindings</span></span><br><span data-ttu-id="17694-653">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-653">
         - TableCoercion</span></span><br><span data-ttu-id="17694-654">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-654">
         - TextBindings</span></span><br><span data-ttu-id="17694-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-655">
         - TextCoercion</span></span><br><span data-ttu-id="17694-656">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-656">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-657">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-657">Office 2019 for Mac</span></span><br><span data-ttu-id="17694-658">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-658">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-659">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="17694-659">- TaskPane</span></span><br><span data-ttu-id="17694-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17694-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="17694-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17694-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="17694-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="17694-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="17694-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-665">- BindingEvents</span></span><br><span data-ttu-id="17694-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-666">
         - CompressedFile</span></span><br><span data-ttu-id="17694-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-668">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-669">
         - File</span></span><br><span data-ttu-id="17694-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-671">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-671">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-672">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-672">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-673">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-673">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-674">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-674">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-675">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-675">
         - PdfFile</span></span><br><span data-ttu-id="17694-676">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-676">
         - Selection</span></span><br><span data-ttu-id="17694-677">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-677">
         - Settings</span></span><br><span data-ttu-id="17694-678">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-678">
         - TableBindings</span></span><br><span data-ttu-id="17694-679">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-679">
         - TableCoercion</span></span><br><span data-ttu-id="17694-680">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-680">
         - TextBindings</span></span><br><span data-ttu-id="17694-681">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-681">
         - TextCoercion</span></span><br><span data-ttu-id="17694-682">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-682">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-683">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-683">Office 2016 for Mac</span></span><br><span data-ttu-id="17694-684">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-684">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-685">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="17694-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17694-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="17694-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17694-688">- BindingEvents</span></span><br><span data-ttu-id="17694-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-689">
         - CompressedFile</span></span><br><span data-ttu-id="17694-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17694-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="17694-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-691">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-692">
         - File</span></span><br><span data-ttu-id="17694-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-694">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-694">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-695">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17694-695">
         - MatrixBindings</span></span><br><span data-ttu-id="17694-696">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-696">
         - MatrixCoercion</span></span><br><span data-ttu-id="17694-697">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-697">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17694-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-698">
         - PdfFile</span></span><br><span data-ttu-id="17694-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-699">
         - Selection</span></span><br><span data-ttu-id="17694-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-700">
         - Settings</span></span><br><span data-ttu-id="17694-701">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17694-701">
         - TableBindings</span></span><br><span data-ttu-id="17694-702">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-702">
         - TableCoercion</span></span><br><span data-ttu-id="17694-703">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17694-703">
         - TextBindings</span></span><br><span data-ttu-id="17694-704">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-704">
         - TextCoercion</span></span><br><span data-ttu-id="17694-705">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17694-705">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="17694-706">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="17694-706">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="17694-707">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="17694-707">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17694-708">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="17694-708">Platform</span></span></th>
    <th><span data-ttu-id="17694-709">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-709">Extension points</span></span></th>
    <th><span data-ttu-id="17694-710">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-710">API requirement sets</span></span></th>
    <th><span data-ttu-id="17694-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-712">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-712">Office Online</span></span></td>
    <td> <span data-ttu-id="17694-713">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-713">- Content</span></span><br><span data-ttu-id="17694-714">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-714">
         - TaskPane</span></span><br><span data-ttu-id="17694-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-717">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-717">- ActiveView</span></span><br><span data-ttu-id="17694-718">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-718">
         - CompressedFile</span></span><br><span data-ttu-id="17694-719">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-719">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-720">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-720">
         - File</span></span><br><span data-ttu-id="17694-721">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-721">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-722">
         - PdfFile</span></span><br><span data-ttu-id="17694-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-723">
         - Selection</span></span><br><span data-ttu-id="17694-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-724">
         - Settings</span></span><br><span data-ttu-id="17694-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-725">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-726">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="17694-726">Office apps on Windows</span></span><br><span data-ttu-id="17694-727">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-727">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-728">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-728">- Content</span></span><br><span data-ttu-id="17694-729">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-729">
         - TaskPane</span></span><br><span data-ttu-id="17694-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-732">- ActiveView</span></span><br><span data-ttu-id="17694-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-733">
         - CompressedFile</span></span><br><span data-ttu-id="17694-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-734">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-735">
         - File</span></span><br><span data-ttu-id="17694-736">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-736">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-737">
         - PdfFile</span></span><br><span data-ttu-id="17694-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-738">
         - Selection</span></span><br><span data-ttu-id="17694-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-739">
         - Settings</span></span><br><span data-ttu-id="17694-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-740">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-741">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-741">Office 2019 for Windows</span></span><br><span data-ttu-id="17694-742">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-742">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-743">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-743">- Content</span></span><br><span data-ttu-id="17694-744">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-744">
         - TaskPane</span></span><br><span data-ttu-id="17694-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-747">- ActiveView</span></span><br><span data-ttu-id="17694-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-748">
         - CompressedFile</span></span><br><span data-ttu-id="17694-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-749">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-750">
         - File</span></span><br><span data-ttu-id="17694-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-751">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-752">
         - PdfFile</span></span><br><span data-ttu-id="17694-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-753">
         - Selection</span></span><br><span data-ttu-id="17694-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-754">
         - Settings</span></span><br><span data-ttu-id="17694-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-756">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-756">Office 2016 for Windows</span></span><br><span data-ttu-id="17694-757">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-757">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-758">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-758">- Content</span></span><br><span data-ttu-id="17694-759">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-759">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17694-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="17694-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-761">- ActiveView</span></span><br><span data-ttu-id="17694-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-762">
         - CompressedFile</span></span><br><span data-ttu-id="17694-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-763">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-764">
         - File</span></span><br><span data-ttu-id="17694-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-765">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-766">
         - PdfFile</span></span><br><span data-ttu-id="17694-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-767">
         - Selection</span></span><br><span data-ttu-id="17694-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-768">
         - Settings</span></span><br><span data-ttu-id="17694-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-770">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-770">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="17694-771">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-772">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-772">- Content</span></span><br><span data-ttu-id="17694-773">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-773">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="17694-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17694-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="17694-775">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-775">- ActiveView</span></span><br><span data-ttu-id="17694-776">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-776">
         - CompressedFile</span></span><br><span data-ttu-id="17694-777">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-777">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-778">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-778">
         - File</span></span><br><span data-ttu-id="17694-779">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-779">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-780">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-780">
         - PdfFile</span></span><br><span data-ttu-id="17694-781">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-781">
         - Selection</span></span><br><span data-ttu-id="17694-782">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-782">
         - Settings</span></span><br><span data-ttu-id="17694-783">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-783">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-784">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="17694-784">Office for iPad</span></span><br><span data-ttu-id="17694-785">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-785">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-786">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-786">- Content</span></span><br><span data-ttu-id="17694-787">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-787">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-789">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-789">- ActiveView</span></span><br><span data-ttu-id="17694-790">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-790">
         - CompressedFile</span></span><br><span data-ttu-id="17694-791">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-791">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-792">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-792">
         - File</span></span><br><span data-ttu-id="17694-793">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-793">
         - PdfFile</span></span><br><span data-ttu-id="17694-794">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-794">
         - Selection</span></span><br><span data-ttu-id="17694-795">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-795">
         - Settings</span></span><br><span data-ttu-id="17694-796">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-796">
         - TextCoercion</span></span><br><span data-ttu-id="17694-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-797">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-798">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-798">Office for Mac</span></span><br><span data-ttu-id="17694-799">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="17694-799">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="17694-800">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-800">- Content</span></span><br><span data-ttu-id="17694-801">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-801">
         - TaskPane</span></span><br><span data-ttu-id="17694-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-804">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-804">- ActiveView</span></span><br><span data-ttu-id="17694-805">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-805">
         - CompressedFile</span></span><br><span data-ttu-id="17694-806">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-806">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-807">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-807">
         - File</span></span><br><span data-ttu-id="17694-808">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-808">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-809">
         - PdfFile</span></span><br><span data-ttu-id="17694-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-810">
         - Selection</span></span><br><span data-ttu-id="17694-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-811">
         - Settings</span></span><br><span data-ttu-id="17694-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-813">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-813">Office 2019 for Mac</span></span><br><span data-ttu-id="17694-814">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-814">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-815">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-815">- Content</span></span><br><span data-ttu-id="17694-816">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-816">
         - TaskPane</span></span><br><span data-ttu-id="17694-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-819">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-819">- ActiveView</span></span><br><span data-ttu-id="17694-820">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-820">
         - CompressedFile</span></span><br><span data-ttu-id="17694-821">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-821">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-822">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-822">
         - File</span></span><br><span data-ttu-id="17694-823">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-823">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-824">
         - PdfFile</span></span><br><span data-ttu-id="17694-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-825">
         - Selection</span></span><br><span data-ttu-id="17694-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-826">
         - Settings</span></span><br><span data-ttu-id="17694-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-828">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-828">Office 2016 for Mac</span></span><br><span data-ttu-id="17694-829">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-829">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-830">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-830">- Content</span></span><br><span data-ttu-id="17694-831">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-831">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17694-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="17694-833">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17694-833">- ActiveView</span></span><br><span data-ttu-id="17694-834">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17694-834">
         - CompressedFile</span></span><br><span data-ttu-id="17694-835">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-835">
         - DocumentEvents</span></span><br><span data-ttu-id="17694-836">
         - File</span><span class="sxs-lookup"><span data-stu-id="17694-836">
         - File</span></span><br><span data-ttu-id="17694-837">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-837">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-838">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17694-838">
         - PdfFile</span></span><br><span data-ttu-id="17694-839">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17694-839">
         - Selection</span></span><br><span data-ttu-id="17694-840">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-840">
         - Settings</span></span><br><span data-ttu-id="17694-841">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-841">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="17694-842">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="17694-842">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="17694-843">OneNote</span><span class="sxs-lookup"><span data-stu-id="17694-843">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17694-844">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="17694-844">Platform</span></span></th>
    <th><span data-ttu-id="17694-845">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-845">Extension points</span></span></th>
    <th><span data-ttu-id="17694-846">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-846">API requirement sets</span></span></th>
    <th><span data-ttu-id="17694-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-848">Office Online</span><span class="sxs-lookup"><span data-stu-id="17694-848">Office Online</span></span></td>
    <td> <span data-ttu-id="17694-849">- Contenu</span><span class="sxs-lookup"><span data-stu-id="17694-849">- Content</span></span><br><span data-ttu-id="17694-850">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-850">
         - TaskPane</span></span><br><span data-ttu-id="17694-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="17694-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17694-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="17694-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-854">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17694-854">- DocumentEvents</span></span><br><span data-ttu-id="17694-855">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-855">
         - HtmlCoercion</span></span><br><span data-ttu-id="17694-856">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-856">
         - ImageCoercion</span></span><br><span data-ttu-id="17694-857">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17694-857">
         - Settings</span></span><br><span data-ttu-id="17694-858">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-858">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="17694-859">Projet</span><span class="sxs-lookup"><span data-stu-id="17694-859">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17694-860">Plateforme</span><span class="sxs-lookup"><span data-stu-id="17694-860">Platform</span></span></th>
    <th><span data-ttu-id="17694-861">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="17694-861">Extension points</span></span></th>
    <th><span data-ttu-id="17694-862">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="17694-862">API requirement sets</span></span></th>
    <th><span data-ttu-id="17694-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="17694-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-864">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-864">Office 2019 for Windows</span></span><br><span data-ttu-id="17694-865">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-866">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-866">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-868">- Selection</span><span class="sxs-lookup"><span data-stu-id="17694-868">- Selection</span></span><br><span data-ttu-id="17694-869">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-869">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-870">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-870">Office 2016 for Windows</span></span><br><span data-ttu-id="17694-871">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-871">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-872">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-872">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-874">- Selection</span><span class="sxs-lookup"><span data-stu-id="17694-874">- Selection</span></span><br><span data-ttu-id="17694-875">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-875">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17694-876">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="17694-876">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="17694-877">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="17694-877">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17694-878">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="17694-878">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17694-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17694-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17694-880">- Selection</span><span class="sxs-lookup"><span data-stu-id="17694-880">- Selection</span></span><br><span data-ttu-id="17694-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17694-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="17694-882">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="17694-882">See also</span></span>

- [<span data-ttu-id="17694-883">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="17694-883">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="17694-884">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="17694-884">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="17694-885">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="17694-885">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="17694-886">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="17694-886">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="17694-887">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="17694-887">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="17694-888">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="17694-888">Update history for Office 365 ProPlus releases</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="17694-889">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="17694-889">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="17694-890">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="17694-890">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="17694-891">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="17694-891">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="17694-892">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="17694-892">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="17694-893">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="17694-893">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
