---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 05/23/2019
localization_priority: Priority
ms.openlocfilehash: 6fb1f0db839910e91d7a5215f8e21f5b33ff2165
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432193"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7b9c9-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="7b9c9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7b9c9-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="7b9c9-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7b9c9-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="7b9c9-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="7b9c9-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="7b9c9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="7b9c9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7b9c9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7b9c9-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7b9c9-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7b9c9-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7b9c9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-113">Office Online</span></span></td>
    <td> <span data-ttu-id="7b9c9-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-114">- TaskPane</span></span><br><span data-ttu-id="7b9c9-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-115">
        - Content</span></span><br><span data-ttu-id="7b9c9-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-116">
        - Custom Functions</span></span><br><span data-ttu-id="7b9c9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="7b9c9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7b9c9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7b9c9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-128">
        - BindingEvents</span></span><br><span data-ttu-id="7b9c9-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-129">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-130">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-131">
        - File</span></span><br><span data-ttu-id="7b9c9-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-132">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-134">
        - Selection</span></span><br><span data-ttu-id="7b9c9-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-135">
        - Settings</span></span><br><span data-ttu-id="7b9c9-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-136">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-137">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-138">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-140">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7b9c9-140">Office on Windows</span></span><br><span data-ttu-id="7b9c9-141">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-142">- TaskPane</span></span><br><span data-ttu-id="7b9c9-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-143">
        - Content</span></span><br><span data-ttu-id="7b9c9-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-144">
        - Custom Functions</span></span><br><span data-ttu-id="7b9c9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="7b9c9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7b9c9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7b9c9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-156">
        - BindingEvents</span></span><br><span data-ttu-id="7b9c9-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-157">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-158">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-159">
        - File</span></span><br><span data-ttu-id="7b9c9-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-160">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-162">
        - Selection</span></span><br><span data-ttu-id="7b9c9-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-163">
        - Settings</span></span><br><span data-ttu-id="7b9c9-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-164">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-165">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-166">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-168">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-168">Office 2019 on Windows</span></span><br><span data-ttu-id="7b9c9-169">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7b9c9-170">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-170">- TaskPane</span></span><br><span data-ttu-id="7b9c9-171">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-171">
        - Content</span></span><br><span data-ttu-id="7b9c9-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7b9c9-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-182">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-183">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-184">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-185">
        - File</span></span><br><span data-ttu-id="7b9c9-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-186">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-187">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-189">
        - Selection</span></span><br><span data-ttu-id="7b9c9-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-190">
        - Settings</span></span><br><span data-ttu-id="7b9c9-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-191">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-192">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-193">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-195">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-195">Office 2016 on Windows</span></span><br><span data-ttu-id="7b9c9-196">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7b9c9-197">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-197">- TaskPane</span></span><br><span data-ttu-id="7b9c9-198">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-198">
        - Content</span></span></td>
    <td><span data-ttu-id="7b9c9-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="7b9c9-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-201">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-202">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-203">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-204">
        - File</span></span><br><span data-ttu-id="7b9c9-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-205">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-206">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-208">
        - Selection</span></span><br><span data-ttu-id="7b9c9-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-209">
        - Settings</span></span><br><span data-ttu-id="7b9c9-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-210">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-211">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-212">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-214">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-214">Office 2013 on Windows</span></span><br><span data-ttu-id="7b9c9-215">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7b9c9-216">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-216">
        - TaskPane</span></span><br><span data-ttu-id="7b9c9-217">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7b9c9-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="7b9c9-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-219">
        - BindingEvents</span></span><br><span data-ttu-id="7b9c9-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-220">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-221">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-222">
        - File</span></span><br><span data-ttu-id="7b9c9-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-223">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-224">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-226">
        - Selection</span></span><br><span data-ttu-id="7b9c9-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-227">
        - Settings</span></span><br><span data-ttu-id="7b9c9-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-228">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-229">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-230">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-232">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="7b9c9-232">Office for iPad</span></span><br><span data-ttu-id="7b9c9-233">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7b9c9-234">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-234">- TaskPane</span></span><br><span data-ttu-id="7b9c9-235">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-235">
        - Content</span></span><br><span data-ttu-id="7b9c9-236">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7b9c9-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7b9c9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-247">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-248">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-249">
        - File</span></span><br><span data-ttu-id="7b9c9-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-250">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-251">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-253">
        - Selection</span></span><br><span data-ttu-id="7b9c9-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-254">
        - Settings</span></span><br><span data-ttu-id="7b9c9-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-255">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-256">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-257">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-259">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-259">Office for Mac</span></span><br><span data-ttu-id="7b9c9-260">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7b9c9-261">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-261">- TaskPane</span></span><br><span data-ttu-id="7b9c9-262">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-262">
        - Content</span></span><br><span data-ttu-id="7b9c9-263">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-263">
        - Custom Functions</span></span><br><span data-ttu-id="7b9c9-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7b9c9-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7b9c9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-275">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-276">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-277">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-278">
        - File</span></span><br><span data-ttu-id="7b9c9-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-279">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-280">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-282">
        - PdfFile</span></span><br><span data-ttu-id="7b9c9-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-283">
        - Selection</span></span><br><span data-ttu-id="7b9c9-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-284">
        - Settings</span></span><br><span data-ttu-id="7b9c9-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-285">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-286">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-287">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-289">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-289">Office 2019 for Mac</span></span><br><span data-ttu-id="7b9c9-290">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7b9c9-291">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-291">- TaskPane</span></span><br><span data-ttu-id="7b9c9-292">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-292">
        - Content</span></span><br><span data-ttu-id="7b9c9-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7b9c9-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7b9c9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7b9c9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7b9c9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7b9c9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7b9c9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7b9c9-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-303">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-304">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-305">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-306">
        - File</span></span><br><span data-ttu-id="7b9c9-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-307">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-308">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-310">
        - PdfFile</span></span><br><span data-ttu-id="7b9c9-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-311">
        - Selection</span></span><br><span data-ttu-id="7b9c9-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-312">
        - Settings</span></span><br><span data-ttu-id="7b9c9-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-313">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-314">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-315">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-317">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-317">Office 2016 for Mac</span></span><br><span data-ttu-id="7b9c9-318">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7b9c9-319">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-319">- TaskPane</span></span><br><span data-ttu-id="7b9c9-320">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-320">
        - Content</span></span></td>
    <td><span data-ttu-id="7b9c9-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="7b9c9-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-323">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-324">
        - CompressedFile</span></span><br><span data-ttu-id="7b9c9-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-325">
        - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-326">
        - File</span></span><br><span data-ttu-id="7b9c9-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-327">
        - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-328">
        - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-330">
        - PdfFile</span></span><br><span data-ttu-id="7b9c9-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-331">
        - Selection</span></span><br><span data-ttu-id="7b9c9-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-332">
        - Settings</span></span><br><span data-ttu-id="7b9c9-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-333">
        - TableBindings</span></span><br><span data-ttu-id="7b9c9-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-334">
        - TableCoercion</span></span><br><span data-ttu-id="7b9c9-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-335">
        - TextBindings</span></span><br><span data-ttu-id="7b9c9-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7b9c9-337">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="7b9c9-338">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7b9c9-339">Plateforme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7b9c9-340">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7b9c9-341">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7b9c9-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-343">Office Online</span></span></td>
    <td><span data-ttu-id="7b9c9-344">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7b9c9-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-346">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7b9c9-346">Office on Windows</span></span><br><span data-ttu-id="7b9c9-347">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7b9c9-348">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7b9c9-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-350">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="7b9c9-350">Office for iPad</span></span><br><span data-ttu-id="7b9c9-351">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7b9c9-352">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7b9c9-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-354">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-354">Office for Mac</span></span><br><span data-ttu-id="7b9c9-355">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7b9c9-356">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7b9c9-356">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7b9c9-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="7b9c9-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="7b9c9-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7b9c9-359">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-359">Platform</span></span></th>
    <th><span data-ttu-id="7b9c9-360">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-360">Extension points</span></span></th>
    <th><span data-ttu-id="7b9c9-361">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="7b9c9-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-363">Office Online</span></span></td>
    <td> <span data-ttu-id="7b9c9-364">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-364">- Mail Read</span></span><br><span data-ttu-id="7b9c9-365">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-365">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7b9c9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7b9c9-374">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-375">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7b9c9-375">Office on Windows</span></span><br><span data-ttu-id="7b9c9-376">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-377">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-377">- Mail Read</span></span><br><span data-ttu-id="7b9c9-378">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-378">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7b9c9-380">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7b9c9-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7b9c9-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7b9c9-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7b9c9-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-389">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-389">Office 2019 on Windows</span></span><br><span data-ttu-id="7b9c9-390">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-391">- Mail Read</span></span><br><span data-ttu-id="7b9c9-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-392">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7b9c9-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7b9c9-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7b9c9-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7b9c9-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7b9c9-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-403">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-403">Office 2016 on Windows</span></span><br><span data-ttu-id="7b9c9-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-405">- Mail Read</span></span><br><span data-ttu-id="7b9c9-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-406">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7b9c9-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7b9c9-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7b9c9-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7b9c9-413">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-414">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-414">Office 2013 on Windows</span></span><br><span data-ttu-id="7b9c9-415">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-416">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-416">- Mail Read</span></span><br><span data-ttu-id="7b9c9-417">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="7b9c9-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="7b9c9-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7b9c9-422">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-423">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="7b9c9-423">Office for iOS</span></span><br><span data-ttu-id="7b9c9-424">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-425">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-425">- Mail Read</span></span><br><span data-ttu-id="7b9c9-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7b9c9-432">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-433">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-433">Office for Mac</span></span><br><span data-ttu-id="7b9c9-434">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-435">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-435">- Mail Read</span></span><br><span data-ttu-id="7b9c9-436">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-436">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7b9c9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7b9c9-445">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-446">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-446">Office 2019 for Mac</span></span><br><span data-ttu-id="7b9c9-447">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-448">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-448">- Mail Read</span></span><br><span data-ttu-id="7b9c9-449">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-449">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7b9c9-457">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-458">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-458">Office 2016 for Mac</span></span><br><span data-ttu-id="7b9c9-459">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-460">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-460">- Mail Read</span></span><br><span data-ttu-id="7b9c9-461">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-461">
      - Mail Compose</span></span><br><span data-ttu-id="7b9c9-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7b9c9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7b9c9-469">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-470">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="7b9c9-470">Office for Android</span></span><br><span data-ttu-id="7b9c9-471">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-471">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-472">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7b9c9-472">- Mail Read</span></span><br><span data-ttu-id="7b9c9-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7b9c9-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7b9c9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7b9c9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7b9c9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7b9c9-479">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7b9c9-479">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="7b9c9-480">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-480">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="7b9c9-481">Word</span><span class="sxs-lookup"><span data-stu-id="7b9c9-481">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7b9c9-482">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-482">Platform</span></span></th>
    <th><span data-ttu-id="7b9c9-483">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-483">Extension points</span></span></th>
    <th><span data-ttu-id="7b9c9-484">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-484">API requirement sets</span></span></th>
    <th><span data-ttu-id="7b9c9-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-486">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-486">Office Online</span></span></td>
    <td> <span data-ttu-id="7b9c9-487">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-487">- TaskPane</span></span><br><span data-ttu-id="7b9c9-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-493">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-493">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-494">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-494">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-495">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-495">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-496">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-496">
         - File</span></span><br><span data-ttu-id="7b9c9-497">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-497">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-498">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-498">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-499">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-499">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-500">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-500">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-501">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-501">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-502">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-502">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-503">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-503">
         - Selection</span></span><br><span data-ttu-id="7b9c9-504">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-504">
         - Settings</span></span><br><span data-ttu-id="7b9c9-505">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-505">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-506">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-506">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-507">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-507">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-508">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-508">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-509">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-509">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-510">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7b9c9-510">Office on Windows</span></span><br><span data-ttu-id="7b9c9-511">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-511">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-512">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-512">- TaskPane</span></span><br><span data-ttu-id="7b9c9-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-518">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-518">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-519">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-519">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-520">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-520">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-521">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-521">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-522">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-522">
         - File</span></span><br><span data-ttu-id="7b9c9-523">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-523">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-524">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-524">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-525">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-525">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-526">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-526">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-527">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-527">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-528">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-528">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-529">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-529">
         - Selection</span></span><br><span data-ttu-id="7b9c9-530">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-530">
         - Settings</span></span><br><span data-ttu-id="7b9c9-531">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-531">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-532">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-532">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-533">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-533">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-534">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-534">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-535">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-535">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-536">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-536">Office 2019 on Windows</span></span><br><span data-ttu-id="7b9c9-537">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-537">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-538">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="7b9c9-538">- TaskPane</span></span><br><span data-ttu-id="7b9c9-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-544">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-544">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-545">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-545">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-546">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-546">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-547">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-547">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-548">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-548">
         - File</span></span><br><span data-ttu-id="7b9c9-549">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-549">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-550">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-550">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-551">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-551">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-552">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-552">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-553">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-553">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-554">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-554">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-555">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-555">
         - Selection</span></span><br><span data-ttu-id="7b9c9-556">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-556">
         - Settings</span></span><br><span data-ttu-id="7b9c9-557">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-557">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-558">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-558">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-559">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-559">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-560">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-561">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-561">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-562">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-562">Office 2016 on Windows</span></span><br><span data-ttu-id="7b9c9-563">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-563">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-564">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-564">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="7b9c9-567">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-567">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-568">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-569">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-569">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-570">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-570">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-571">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-571">
         - File</span></span><br><span data-ttu-id="7b9c9-572">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-572">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-573">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-574">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-577">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-578">
         - Selection</span></span><br><span data-ttu-id="7b9c9-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-579">
         - Settings</span></span><br><span data-ttu-id="7b9c9-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-580">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-581">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-582">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-583">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-585">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-585">Office 2013 on Windows</span></span><br><span data-ttu-id="7b9c9-586">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-587">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7b9c9-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-589">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-590">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-592">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-593">
         - File</span></span><br><span data-ttu-id="7b9c9-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-595">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-596">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-599">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-600">
         - Selection</span></span><br><span data-ttu-id="7b9c9-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-601">
         - Settings</span></span><br><span data-ttu-id="7b9c9-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-602">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-603">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-604">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-605">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-606">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-607">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="7b9c9-607">Office for iPad</span></span><br><span data-ttu-id="7b9c9-608">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-608">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-609">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7b9c9-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7b9c9-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-614">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-615">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-617">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-618">
         - File</span></span><br><span data-ttu-id="7b9c9-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-620">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-621">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-624">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-625">
         - Selection</span></span><br><span data-ttu-id="7b9c9-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-626">
         - Settings</span></span><br><span data-ttu-id="7b9c9-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-627">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-628">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-629">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-630">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-632">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-632">Office for Mac</span></span><br><span data-ttu-id="7b9c9-633">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-633">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-634">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-634">- TaskPane</span></span><br><span data-ttu-id="7b9c9-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7b9c9-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7b9c9-640">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-640">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-641">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-641">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-642">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-642">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-643">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-643">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-644">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-644">
         - File</span></span><br><span data-ttu-id="7b9c9-645">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-645">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-646">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-646">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-647">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-647">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-648">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-648">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-649">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-649">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-650">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-651">
         - Selection</span></span><br><span data-ttu-id="7b9c9-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-652">
         - Settings</span></span><br><span data-ttu-id="7b9c9-653">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-653">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-654">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-654">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-655">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-655">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-656">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-657">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-657">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-658">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-658">Office 2019 for Mac</span></span><br><span data-ttu-id="7b9c9-659">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-659">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-660">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="7b9c9-660">- TaskPane</span></span><br><span data-ttu-id="7b9c9-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7b9c9-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7b9c9-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7b9c9-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7b9c9-666">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-666">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-667">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-668">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-668">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-669">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-669">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-670">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-670">
         - File</span></span><br><span data-ttu-id="7b9c9-671">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-671">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-672">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-672">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-673">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-673">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-674">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-674">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-675">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-675">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-676">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-676">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-677">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-677">
         - Selection</span></span><br><span data-ttu-id="7b9c9-678">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-678">
         - Settings</span></span><br><span data-ttu-id="7b9c9-679">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-679">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-680">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-680">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-681">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-681">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-682">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-682">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-683">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-683">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-684">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-684">Office 2016 for Mac</span></span><br><span data-ttu-id="7b9c9-685">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-685">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-686">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="7b9c9-689">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-689">- BindingEvents</span></span><br><span data-ttu-id="7b9c9-690">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-690">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-691">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7b9c9-691">
         - CustomXmlParts</span></span><br><span data-ttu-id="7b9c9-692">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-692">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-693">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-693">
         - File</span></span><br><span data-ttu-id="7b9c9-694">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-694">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-695">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-695">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-696">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-696">
         - MatrixBindings</span></span><br><span data-ttu-id="7b9c9-697">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-697">
         - MatrixCoercion</span></span><br><span data-ttu-id="7b9c9-698">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-698">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7b9c9-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-699">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-700">
         - Selection</span></span><br><span data-ttu-id="7b9c9-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-701">
         - Settings</span></span><br><span data-ttu-id="7b9c9-702">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-702">
         - TableBindings</span></span><br><span data-ttu-id="7b9c9-703">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-703">
         - TableCoercion</span></span><br><span data-ttu-id="7b9c9-704">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-704">
         - TextBindings</span></span><br><span data-ttu-id="7b9c9-705">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-705">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-706">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-706">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="7b9c9-707">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-707">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7b9c9-708">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7b9c9-708">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7b9c9-709">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-709">Platform</span></span></th>
    <th><span data-ttu-id="7b9c9-710">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-710">Extension points</span></span></th>
    <th><span data-ttu-id="7b9c9-711">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-711">API requirement sets</span></span></th>
    <th><span data-ttu-id="7b9c9-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-713">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-713">Office Online</span></span></td>
    <td> <span data-ttu-id="7b9c9-714">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-714">- Content</span></span><br><span data-ttu-id="7b9c9-715">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-715">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-718">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-718">- ActiveView</span></span><br><span data-ttu-id="7b9c9-719">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-719">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-720">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-720">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-721">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-721">
         - File</span></span><br><span data-ttu-id="7b9c9-722">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-722">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-723">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-724">
         - Selection</span></span><br><span data-ttu-id="7b9c9-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-725">
         - Settings</span></span><br><span data-ttu-id="7b9c9-726">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-726">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-727">Office sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7b9c9-727">Office on Windows</span></span><br><span data-ttu-id="7b9c9-728">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-728">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-729">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-729">- Content</span></span><br><span data-ttu-id="7b9c9-730">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-730">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-733">- ActiveView</span></span><br><span data-ttu-id="7b9c9-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-734">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-735">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-736">
         - File</span></span><br><span data-ttu-id="7b9c9-737">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-737">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-738">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-738">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-739">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-739">
         - Selection</span></span><br><span data-ttu-id="7b9c9-740">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-740">
         - Settings</span></span><br><span data-ttu-id="7b9c9-741">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-741">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-742">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-742">Office 2019 on Windows</span></span><br><span data-ttu-id="7b9c9-743">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-743">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-744">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-744">- Content</span></span><br><span data-ttu-id="7b9c9-745">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-745">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-748">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-748">- ActiveView</span></span><br><span data-ttu-id="7b9c9-749">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-749">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-750">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-750">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-751">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-751">
         - File</span></span><br><span data-ttu-id="7b9c9-752">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-752">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-753">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-753">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-754">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-754">
         - Selection</span></span><br><span data-ttu-id="7b9c9-755">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-755">
         - Settings</span></span><br><span data-ttu-id="7b9c9-756">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-756">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-757">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-757">Office 2016 on Windows</span></span><br><span data-ttu-id="7b9c9-758">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-758">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-759">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-759">- Content</span></span><br><span data-ttu-id="7b9c9-760">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-760">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7b9c9-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-762">- ActiveView</span></span><br><span data-ttu-id="7b9c9-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-763">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-764">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-765">
         - File</span></span><br><span data-ttu-id="7b9c9-766">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-766">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-767">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-768">
         - Selection</span></span><br><span data-ttu-id="7b9c9-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-769">
         - Settings</span></span><br><span data-ttu-id="7b9c9-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-771">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-771">Office 2013 on Windows</span></span><br><span data-ttu-id="7b9c9-772">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-772">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-773">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-773">- Content</span></span><br><span data-ttu-id="7b9c9-774">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-774">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7b9c9-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7b9c9-776">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-776">- ActiveView</span></span><br><span data-ttu-id="7b9c9-777">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-777">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-778">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-778">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-779">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-779">
         - File</span></span><br><span data-ttu-id="7b9c9-780">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-780">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-781">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-782">
         - Selection</span></span><br><span data-ttu-id="7b9c9-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-783">
         - Settings</span></span><br><span data-ttu-id="7b9c9-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-785">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="7b9c9-785">Office for iPad</span></span><br><span data-ttu-id="7b9c9-786">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-786">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-787">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-787">- Content</span></span><br><span data-ttu-id="7b9c9-788">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-790">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-790">- ActiveView</span></span><br><span data-ttu-id="7b9c9-791">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-791">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-792">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-792">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-793">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-793">
         - File</span></span><br><span data-ttu-id="7b9c9-794">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-794">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-795">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-795">
         - Selection</span></span><br><span data-ttu-id="7b9c9-796">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-796">
         - Settings</span></span><br><span data-ttu-id="7b9c9-797">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-797">
         - TextCoercion</span></span><br><span data-ttu-id="7b9c9-798">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-798">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-799">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-799">Office for Mac</span></span><br><span data-ttu-id="7b9c9-800">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-800">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="7b9c9-801">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-801">- Content</span></span><br><span data-ttu-id="7b9c9-802">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-802">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-805">- ActiveView</span></span><br><span data-ttu-id="7b9c9-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-806">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-807">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-808">
         - File</span></span><br><span data-ttu-id="7b9c9-809">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-809">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-810">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-810">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-811">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-811">
         - Selection</span></span><br><span data-ttu-id="7b9c9-812">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-812">
         - Settings</span></span><br><span data-ttu-id="7b9c9-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-814">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-814">Office 2019 for Mac</span></span><br><span data-ttu-id="7b9c9-815">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-815">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-816">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-816">- Content</span></span><br><span data-ttu-id="7b9c9-817">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-817">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-820">- ActiveView</span></span><br><span data-ttu-id="7b9c9-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-821">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-822">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-823">
         - File</span></span><br><span data-ttu-id="7b9c9-824">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-824">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-825">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-825">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-826">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-826">
         - Selection</span></span><br><span data-ttu-id="7b9c9-827">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-827">
         - Settings</span></span><br><span data-ttu-id="7b9c9-828">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-828">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-829">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-829">Office 2016 for Mac</span></span><br><span data-ttu-id="7b9c9-830">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-830">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-831">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-831">- Content</span></span><br><span data-ttu-id="7b9c9-832">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-832">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7b9c9-834">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7b9c9-834">- ActiveView</span></span><br><span data-ttu-id="7b9c9-835">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-835">
         - CompressedFile</span></span><br><span data-ttu-id="7b9c9-836">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-836">
         - DocumentEvents</span></span><br><span data-ttu-id="7b9c9-837">
         - File</span><span class="sxs-lookup"><span data-stu-id="7b9c9-837">
         - File</span></span><br><span data-ttu-id="7b9c9-838">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-838">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-839">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7b9c9-839">
         - PdfFile</span></span><br><span data-ttu-id="7b9c9-840">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-840">
         - Selection</span></span><br><span data-ttu-id="7b9c9-841">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-841">
         - Settings</span></span><br><span data-ttu-id="7b9c9-842">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-842">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7b9c9-843">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7b9c9-843">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="7b9c9-844">OneNote</span><span class="sxs-lookup"><span data-stu-id="7b9c9-844">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7b9c9-845">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-845">Platform</span></span></th>
    <th><span data-ttu-id="7b9c9-846">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-846">Extension points</span></span></th>
    <th><span data-ttu-id="7b9c9-847">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-847">API requirement sets</span></span></th>
    <th><span data-ttu-id="7b9c9-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-849">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b9c9-849">Office Online</span></span></td>
    <td> <span data-ttu-id="7b9c9-850">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7b9c9-850">- Content</span></span><br><span data-ttu-id="7b9c9-851">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-851">
         - TaskPane</span></span><br><span data-ttu-id="7b9c9-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7b9c9-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-855">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7b9c9-855">- DocumentEvents</span></span><br><span data-ttu-id="7b9c9-856">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-856">
         - HtmlCoercion</span></span><br><span data-ttu-id="7b9c9-857">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-857">
         - ImageCoercion</span></span><br><span data-ttu-id="7b9c9-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7b9c9-858">
         - Settings</span></span><br><span data-ttu-id="7b9c9-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-859">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7b9c9-860">Projet</span><span class="sxs-lookup"><span data-stu-id="7b9c9-860">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7b9c9-861">Plateforme</span><span class="sxs-lookup"><span data-stu-id="7b9c9-861">Platform</span></span></th>
    <th><span data-ttu-id="7b9c9-862">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7b9c9-862">Extension points</span></span></th>
    <th><span data-ttu-id="7b9c9-863">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7b9c9-863">API requirement sets</span></span></th>
    <th><span data-ttu-id="7b9c9-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-865">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-865">Office 2019 on Windows</span></span><br><span data-ttu-id="7b9c9-866">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-866">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-867">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-867">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-869">- Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-869">- Selection</span></span><br><span data-ttu-id="7b9c9-870">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-870">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-871">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-871">Office 2016 on Windows</span></span><br><span data-ttu-id="7b9c9-872">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-872">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-873">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-873">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-875">- Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-875">- Selection</span></span><br><span data-ttu-id="7b9c9-876">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-876">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7b9c9-877">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7b9c9-877">Office 2013 on Windows</span></span><br><span data-ttu-id="7b9c9-878">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-878">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7b9c9-879">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-879">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7b9c9-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7b9c9-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7b9c9-881">- Selection</span><span class="sxs-lookup"><span data-stu-id="7b9c9-881">- Selection</span></span><br><span data-ttu-id="7b9c9-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7b9c9-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7b9c9-883">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7b9c9-883">See also</span></span>

- [<span data-ttu-id="7b9c9-884">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-884">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7b9c9-885">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7b9c9-885">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="7b9c9-886">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="7b9c9-886">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="7b9c9-887">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="7b9c9-887">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="7b9c9-888">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="7b9c9-888">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="7b9c9-889">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="7b9c9-889">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="7b9c9-890">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-890">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="7b9c9-891">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-891">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="7b9c9-892">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-892">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="7b9c9-893">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="7b9c9-893">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="7b9c9-894">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b9c9-894">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
