---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 06/13/2019
localization_priority: Priority
ms.openlocfilehash: 82c276c802cab66ae4f5443d0d556bc42ee57841
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128621"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="708a1-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="708a1-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="708a1-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="708a1-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="708a1-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="708a1-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="708a1-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="708a1-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="708a1-108">Excel</span><span class="sxs-lookup"><span data-stu-id="708a1-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="708a1-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="708a1-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="708a1-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="708a1-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="708a1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="708a1-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-114">- TaskPane</span></span><br><span data-ttu-id="708a1-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-115">
        - Content</span></span><br><span data-ttu-id="708a1-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-116">
        - Custom Functions</span></span><br><span data-ttu-id="708a1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="708a1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="708a1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="708a1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="708a1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-128">
        - BindingEvents</span></span><br><span data-ttu-id="708a1-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-129">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-130">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-131">
        - File</span></span><br><span data-ttu-id="708a1-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-132">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-134">
        - Selection</span></span><br><span data-ttu-id="708a1-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-135">
        - Settings</span></span><br><span data-ttu-id="708a1-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-136">
        - TableBindings</span></span><br><span data-ttu-id="708a1-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-137">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-138">
        - TextBindings</span></span><br><span data-ttu-id="708a1-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-140">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-140">Office on Windows</span></span><br><span data-ttu-id="708a1-141">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-142">- TaskPane</span></span><br><span data-ttu-id="708a1-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-143">
        - Content</span></span><br><span data-ttu-id="708a1-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-144">
        - Custom Functions</span></span><br><span data-ttu-id="708a1-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="708a1-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="708a1-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="708a1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="708a1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-156">
        - BindingEvents</span></span><br><span data-ttu-id="708a1-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-157">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-158">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-159">
        - File</span></span><br><span data-ttu-id="708a1-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-160">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-162">
        - Selection</span></span><br><span data-ttu-id="708a1-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-163">
        - Settings</span></span><br><span data-ttu-id="708a1-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-164">
        - TableBindings</span></span><br><span data-ttu-id="708a1-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-165">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-166">
        - TextBindings</span></span><br><span data-ttu-id="708a1-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-168">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-168">Office 2019 on Windows</span></span><br><span data-ttu-id="708a1-169">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="708a1-170">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-170">- TaskPane</span></span><br><span data-ttu-id="708a1-171">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-171">
        - Content</span></span><br><span data-ttu-id="708a1-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="708a1-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-182">- BindingEvents</span></span><br><span data-ttu-id="708a1-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-183">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-184">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-185">
        - File</span></span><br><span data-ttu-id="708a1-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-186">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-187">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-189">
        - Selection</span></span><br><span data-ttu-id="708a1-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-190">
        - Settings</span></span><br><span data-ttu-id="708a1-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-191">
        - TableBindings</span></span><br><span data-ttu-id="708a1-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-192">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-193">
        - TextBindings</span></span><br><span data-ttu-id="708a1-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-195">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-195">Office 2016 on Windows</span></span><br><span data-ttu-id="708a1-196">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="708a1-197">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-197">- TaskPane</span></span><br><span data-ttu-id="708a1-198">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-198">
        - Content</span></span></td>
    <td><span data-ttu-id="708a1-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="708a1-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-201">- BindingEvents</span></span><br><span data-ttu-id="708a1-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-202">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-203">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-204">
        - File</span></span><br><span data-ttu-id="708a1-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-205">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-206">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-208">
        - Selection</span></span><br><span data-ttu-id="708a1-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-209">
        - Settings</span></span><br><span data-ttu-id="708a1-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-210">
        - TableBindings</span></span><br><span data-ttu-id="708a1-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-211">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-212">
        - TextBindings</span></span><br><span data-ttu-id="708a1-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-214">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-214">Office 2013 on Windows</span></span><br><span data-ttu-id="708a1-215">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="708a1-216">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-216">
        - TaskPane</span></span><br><span data-ttu-id="708a1-217">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="708a1-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="708a1-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="708a1-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-219">
        - BindingEvents</span></span><br><span data-ttu-id="708a1-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-220">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-221">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-222">
        - File</span></span><br><span data-ttu-id="708a1-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-223">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-224">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-226">
        - Selection</span></span><br><span data-ttu-id="708a1-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-227">
        - Settings</span></span><br><span data-ttu-id="708a1-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-228">
        - TableBindings</span></span><br><span data-ttu-id="708a1-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-229">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-230">
        - TextBindings</span></span><br><span data-ttu-id="708a1-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-232">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="708a1-232">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="708a1-233">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-233">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="708a1-234">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-234">- TaskPane</span></span><br><span data-ttu-id="708a1-235">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-235">
        - Content</span></span><br><span data-ttu-id="708a1-236">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="708a1-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="708a1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="708a1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-247">- BindingEvents</span></span><br><span data-ttu-id="708a1-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-248">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-249">
        - File</span></span><br><span data-ttu-id="708a1-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-250">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-251">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-253">
        - Selection</span></span><br><span data-ttu-id="708a1-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-254">
        - Settings</span></span><br><span data-ttu-id="708a1-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-255">
        - TableBindings</span></span><br><span data-ttu-id="708a1-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-256">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-257">
        - TextBindings</span></span><br><span data-ttu-id="708a1-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-259">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-259">Office apps on Mac</span></span><br><span data-ttu-id="708a1-260">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-260">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="708a1-261">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-261">- TaskPane</span></span><br><span data-ttu-id="708a1-262">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-262">
        - Content</span></span><br><span data-ttu-id="708a1-263">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-263">
        - Custom Functions</span></span><br><span data-ttu-id="708a1-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="708a1-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="708a1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="708a1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-275">- BindingEvents</span></span><br><span data-ttu-id="708a1-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-276">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-277">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-278">
        - File</span></span><br><span data-ttu-id="708a1-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-279">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-280">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-282">
        - PdfFile</span></span><br><span data-ttu-id="708a1-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-283">
        - Selection</span></span><br><span data-ttu-id="708a1-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-284">
        - Settings</span></span><br><span data-ttu-id="708a1-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-285">
        - TableBindings</span></span><br><span data-ttu-id="708a1-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-286">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-287">
        - TextBindings</span></span><br><span data-ttu-id="708a1-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-289">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-289">Office 2019 for Mac</span></span><br><span data-ttu-id="708a1-290">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="708a1-291">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-291">- TaskPane</span></span><br><span data-ttu-id="708a1-292">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-292">
        - Content</span></span><br><span data-ttu-id="708a1-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="708a1-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="708a1-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="708a1-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="708a1-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="708a1-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="708a1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="708a1-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="708a1-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="708a1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="708a1-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-303">- BindingEvents</span></span><br><span data-ttu-id="708a1-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-304">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-305">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-306">
        - File</span></span><br><span data-ttu-id="708a1-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-307">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-308">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-310">
        - PdfFile</span></span><br><span data-ttu-id="708a1-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-311">
        - Selection</span></span><br><span data-ttu-id="708a1-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-312">
        - Settings</span></span><br><span data-ttu-id="708a1-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-313">
        - TableBindings</span></span><br><span data-ttu-id="708a1-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-314">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-315">
        - TextBindings</span></span><br><span data-ttu-id="708a1-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-317">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-317">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="708a1-318">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="708a1-319">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-319">- TaskPane</span></span><br><span data-ttu-id="708a1-320">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-320">
        - Content</span></span></td>
    <td><span data-ttu-id="708a1-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="708a1-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="708a1-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-323">- BindingEvents</span></span><br><span data-ttu-id="708a1-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-324">
        - CompressedFile</span></span><br><span data-ttu-id="708a1-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-325">
        - DocumentEvents</span></span><br><span data-ttu-id="708a1-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="708a1-326">
        - File</span></span><br><span data-ttu-id="708a1-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-327">
        - ImageCoercion</span></span><br><span data-ttu-id="708a1-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-328">
        - MatrixBindings</span></span><br><span data-ttu-id="708a1-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="708a1-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-330">
        - PdfFile</span></span><br><span data-ttu-id="708a1-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-331">
        - Selection</span></span><br><span data-ttu-id="708a1-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-332">
        - Settings</span></span><br><span data-ttu-id="708a1-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-333">
        - TableBindings</span></span><br><span data-ttu-id="708a1-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-334">
        - TableCoercion</span></span><br><span data-ttu-id="708a1-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-335">
        - TextBindings</span></span><br><span data-ttu-id="708a1-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="708a1-337">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="708a1-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="708a1-338">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="708a1-339">Plateforme</span><span class="sxs-lookup"><span data-stu-id="708a1-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="708a1-340">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="708a1-341">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="708a1-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-343">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-343">Office on the web</span></span></td>
    <td><span data-ttu-id="708a1-344">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="708a1-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-346">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-346">Office on Windows</span></span><br><span data-ttu-id="708a1-347">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-347">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="708a1-348">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="708a1-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-350">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-350">Office for Mac</span></span><br><span data-ttu-id="708a1-351">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="708a1-352">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="708a1-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="708a1-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="708a1-354">Outlook</span><span class="sxs-lookup"><span data-stu-id="708a1-354">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="708a1-355">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="708a1-355">Platform</span></span></th>
    <th><span data-ttu-id="708a1-356">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-356">Extension points</span></span></th>
    <th><span data-ttu-id="708a1-357">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="708a1-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-359">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-359">Office on the web</span></span><br><span data-ttu-id="708a1-360">(nouveau)</span><span class="sxs-lookup"><span data-stu-id="708a1-360">New</span></span></td>
    <td> <span data-ttu-id="708a1-361">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-361">- Mail Read</span></span><br><span data-ttu-id="708a1-362">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-362">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="708a1-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="708a1-371">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-372">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-372">Office on the web</span></span><br><span data-ttu-id="708a1-373">(classique)</span><span class="sxs-lookup"><span data-stu-id="708a1-373">Classic.</span></span></td>
    <td> <span data-ttu-id="708a1-374">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-374">- Mail Read</span></span><br><span data-ttu-id="708a1-375">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-375">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="708a1-383">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-384">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-384">Office on Windows</span></span><br><span data-ttu-id="708a1-385">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-385">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-386">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-386">- Mail Read</span></span><br><span data-ttu-id="708a1-387">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-387">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="708a1-389">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="708a1-389">
      - Modules</span></span></td>
    <td> <span data-ttu-id="708a1-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="708a1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="708a1-397">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-398">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-398">Office 2019 on Windows</span></span><br><span data-ttu-id="708a1-399">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-399">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-400">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-400">- Mail Read</span></span><br><span data-ttu-id="708a1-401">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-401">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="708a1-403">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="708a1-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="708a1-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="708a1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="708a1-411">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-412">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-412">Office 2016 on Windows</span></span><br><span data-ttu-id="708a1-413">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-413">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-414">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-414">- Mail Read</span></span><br><span data-ttu-id="708a1-415">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-415">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="708a1-417">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="708a1-417">
      - Modules</span></span></td>
    <td> <span data-ttu-id="708a1-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="708a1-422">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-423">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-423">Office 2013 on Windows</span></span><br><span data-ttu-id="708a1-424">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-424">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-425">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-425">- Mail Read</span></span><br><span data-ttu-id="708a1-426">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-426">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="708a1-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="708a1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="708a1-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-432">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="708a1-432">Office apps on iOS</span></span><br><span data-ttu-id="708a1-433">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-433">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-434">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-434">- Mail Read</span></span><br><span data-ttu-id="708a1-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="708a1-441">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-442">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-442">Office apps on Mac</span></span><br><span data-ttu-id="708a1-443">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-443">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-444">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-444">- Mail Read</span></span><br><span data-ttu-id="708a1-445">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-445">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="708a1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="708a1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="708a1-454">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-455">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-455">Office 2019 for Mac</span></span><br><span data-ttu-id="708a1-456">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-456">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-457">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-457">- Mail Read</span></span><br><span data-ttu-id="708a1-458">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-458">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="708a1-466">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-467">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-467">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="708a1-468">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-468">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-469">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-469">- Mail Read</span></span><br><span data-ttu-id="708a1-470">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="708a1-470">
      - Mail Compose</span></span><br><span data-ttu-id="708a1-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="708a1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="708a1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="708a1-478">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-479">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="708a1-479">Office apps on Android</span></span><br><span data-ttu-id="708a1-480">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-480">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-481">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="708a1-481">- Mail Read</span></span><br><span data-ttu-id="708a1-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="708a1-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="708a1-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="708a1-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="708a1-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="708a1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="708a1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="708a1-488">Non disponible</span><span class="sxs-lookup"><span data-stu-id="708a1-488">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="708a1-489">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="708a1-489">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="708a1-490">Word</span><span class="sxs-lookup"><span data-stu-id="708a1-490">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="708a1-491">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="708a1-491">Platform</span></span></th>
    <th><span data-ttu-id="708a1-492">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-492">Extension points</span></span></th>
    <th><span data-ttu-id="708a1-493">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-493">API requirement sets</span></span></th>
    <th><span data-ttu-id="708a1-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-495">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-495">Office on the web</span></span></td>
    <td> <span data-ttu-id="708a1-496">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-496">- TaskPane</span></span><br><span data-ttu-id="708a1-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-502">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-502">- BindingEvents</span></span><br><span data-ttu-id="708a1-503">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-503">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-504">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-504">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-505">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-505">
         - File</span></span><br><span data-ttu-id="708a1-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-506">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-507">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-508">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-508">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-509">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-509">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-510">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-510">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-511">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-511">
         - PdfFile</span></span><br><span data-ttu-id="708a1-512">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-512">
         - Selection</span></span><br><span data-ttu-id="708a1-513">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-513">
         - Settings</span></span><br><span data-ttu-id="708a1-514">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-514">
         - TableBindings</span></span><br><span data-ttu-id="708a1-515">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-515">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-516">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-516">
         - TextBindings</span></span><br><span data-ttu-id="708a1-517">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-517">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-518">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-518">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-519">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-519">Office on Windows</span></span><br><span data-ttu-id="708a1-520">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-520">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-521">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-521">- TaskPane</span></span><br><span data-ttu-id="708a1-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-527">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-527">- BindingEvents</span></span><br><span data-ttu-id="708a1-528">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-528">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-529">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-529">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-530">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-530">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-531">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-531">
         - File</span></span><br><span data-ttu-id="708a1-532">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-532">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-533">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-533">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-534">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-534">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-535">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-535">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-536">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-536">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-537">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-537">
         - PdfFile</span></span><br><span data-ttu-id="708a1-538">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-538">
         - Selection</span></span><br><span data-ttu-id="708a1-539">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-539">
         - Settings</span></span><br><span data-ttu-id="708a1-540">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-540">
         - TableBindings</span></span><br><span data-ttu-id="708a1-541">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-541">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-542">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-542">
         - TextBindings</span></span><br><span data-ttu-id="708a1-543">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-543">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-544">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-544">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-545">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-545">Office 2019 on Windows</span></span><br><span data-ttu-id="708a1-546">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-546">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-547">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="708a1-547">- TaskPane</span></span><br><span data-ttu-id="708a1-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-553">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-553">- BindingEvents</span></span><br><span data-ttu-id="708a1-554">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-554">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-555">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-555">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-556">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-557">
         - File</span></span><br><span data-ttu-id="708a1-558">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-558">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-559">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-559">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-560">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-560">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-561">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-561">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-562">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-562">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-563">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-563">
         - PdfFile</span></span><br><span data-ttu-id="708a1-564">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-564">
         - Selection</span></span><br><span data-ttu-id="708a1-565">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-565">
         - Settings</span></span><br><span data-ttu-id="708a1-566">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-566">
         - TableBindings</span></span><br><span data-ttu-id="708a1-567">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-567">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-568">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-568">
         - TextBindings</span></span><br><span data-ttu-id="708a1-569">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-569">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-570">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-570">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-571">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-571">Office 2016 on Windows</span></span><br><span data-ttu-id="708a1-572">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-572">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-573">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-573">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="708a1-576">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-576">- BindingEvents</span></span><br><span data-ttu-id="708a1-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-577">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-578">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-578">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-579">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-580">
         - File</span></span><br><span data-ttu-id="708a1-581">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-581">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-582">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-583">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-583">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-584">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-584">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-585">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-585">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-586">
         - PdfFile</span></span><br><span data-ttu-id="708a1-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-587">
         - Selection</span></span><br><span data-ttu-id="708a1-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-588">
         - Settings</span></span><br><span data-ttu-id="708a1-589">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-589">
         - TableBindings</span></span><br><span data-ttu-id="708a1-590">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-590">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-591">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-591">
         - TextBindings</span></span><br><span data-ttu-id="708a1-592">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-592">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-593">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-593">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-594">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-594">Office 2013 on Windows</span></span><br><span data-ttu-id="708a1-595">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-595">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-596">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-596">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="708a1-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="708a1-598">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-598">- BindingEvents</span></span><br><span data-ttu-id="708a1-599">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-599">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-600">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-600">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-601">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-601">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-602">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-602">
         - File</span></span><br><span data-ttu-id="708a1-603">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-603">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-604">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-604">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-605">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-605">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-606">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-606">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-607">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-607">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-608">
         - PdfFile</span></span><br><span data-ttu-id="708a1-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-609">
         - Selection</span></span><br><span data-ttu-id="708a1-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-610">
         - Settings</span></span><br><span data-ttu-id="708a1-611">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-611">
         - TableBindings</span></span><br><span data-ttu-id="708a1-612">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-612">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-613">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-613">
         - TextBindings</span></span><br><span data-ttu-id="708a1-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-614">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-615">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-615">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-616">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="708a1-616">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="708a1-617">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-617">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-618">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-618">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="708a1-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="708a1-623">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-623">- BindingEvents</span></span><br><span data-ttu-id="708a1-624">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-624">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-625">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-625">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-626">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-627">
         - File</span></span><br><span data-ttu-id="708a1-628">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-628">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-629">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-629">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-630">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-630">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-631">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-631">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-632">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-632">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-633">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-633">
         - PdfFile</span></span><br><span data-ttu-id="708a1-634">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-634">
         - Selection</span></span><br><span data-ttu-id="708a1-635">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-635">
         - Settings</span></span><br><span data-ttu-id="708a1-636">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-636">
         - TableBindings</span></span><br><span data-ttu-id="708a1-637">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-637">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-638">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-638">
         - TextBindings</span></span><br><span data-ttu-id="708a1-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-639">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-640">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-640">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-641">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-641">Office apps on Mac</span></span><br><span data-ttu-id="708a1-642">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-642">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-643">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-643">- TaskPane</span></span><br><span data-ttu-id="708a1-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="708a1-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="708a1-649">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-649">- BindingEvents</span></span><br><span data-ttu-id="708a1-650">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-650">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-651">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-651">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-652">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-652">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-653">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-653">
         - File</span></span><br><span data-ttu-id="708a1-654">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-654">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-655">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-655">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-656">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-656">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-657">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-657">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-658">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-658">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-659">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-659">
         - PdfFile</span></span><br><span data-ttu-id="708a1-660">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-660">
         - Selection</span></span><br><span data-ttu-id="708a1-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-661">
         - Settings</span></span><br><span data-ttu-id="708a1-662">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-662">
         - TableBindings</span></span><br><span data-ttu-id="708a1-663">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-663">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-664">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-664">
         - TextBindings</span></span><br><span data-ttu-id="708a1-665">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-665">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-666">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-666">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-667">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-667">Office 2019 for Mac</span></span><br><span data-ttu-id="708a1-668">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-668">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-669">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="708a1-669">- TaskPane</span></span><br><span data-ttu-id="708a1-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="708a1-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="708a1-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="708a1-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="708a1-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="708a1-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="708a1-675">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-675">- BindingEvents</span></span><br><span data-ttu-id="708a1-676">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-676">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-677">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-677">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-678">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-678">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-679">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-679">
         - File</span></span><br><span data-ttu-id="708a1-680">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-680">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-681">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-681">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-682">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-682">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-683">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-683">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-684">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-684">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-685">
         - PdfFile</span></span><br><span data-ttu-id="708a1-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-686">
         - Selection</span></span><br><span data-ttu-id="708a1-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-687">
         - Settings</span></span><br><span data-ttu-id="708a1-688">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-688">
         - TableBindings</span></span><br><span data-ttu-id="708a1-689">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-689">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-690">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-690">
         - TextBindings</span></span><br><span data-ttu-id="708a1-691">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-691">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-692">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-692">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-693">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-693">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="708a1-694">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-694">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-695">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-695">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="708a1-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="708a1-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="708a1-698">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-698">- BindingEvents</span></span><br><span data-ttu-id="708a1-699">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-699">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-700">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="708a1-700">
         - CustomXmlParts</span></span><br><span data-ttu-id="708a1-701">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-701">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-702">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-702">
         - File</span></span><br><span data-ttu-id="708a1-703">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-703">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-704">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-704">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-705">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-705">
         - MatrixBindings</span></span><br><span data-ttu-id="708a1-706">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-706">
         - MatrixCoercion</span></span><br><span data-ttu-id="708a1-707">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-707">
         - OoxmlCoercion</span></span><br><span data-ttu-id="708a1-708">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-708">
         - PdfFile</span></span><br><span data-ttu-id="708a1-709">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-709">
         - Selection</span></span><br><span data-ttu-id="708a1-710">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-710">
         - Settings</span></span><br><span data-ttu-id="708a1-711">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-711">
         - TableBindings</span></span><br><span data-ttu-id="708a1-712">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-712">
         - TableCoercion</span></span><br><span data-ttu-id="708a1-713">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="708a1-713">
         - TextBindings</span></span><br><span data-ttu-id="708a1-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-714">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-715">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="708a1-715">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="708a1-716">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="708a1-716">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="708a1-717">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="708a1-717">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="708a1-718">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="708a1-718">Platform</span></span></th>
    <th><span data-ttu-id="708a1-719">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-719">Extension points</span></span></th>
    <th><span data-ttu-id="708a1-720">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-720">API requirement sets</span></span></th>
    <th><span data-ttu-id="708a1-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-722">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-722">Office on the web</span></span></td>
    <td> <span data-ttu-id="708a1-723">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-723">- Content</span></span><br><span data-ttu-id="708a1-724">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-724">
         - TaskPane</span></span><br><span data-ttu-id="708a1-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-727">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-727">- ActiveView</span></span><br><span data-ttu-id="708a1-728">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-728">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-729">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-729">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-730">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-730">
         - File</span></span><br><span data-ttu-id="708a1-731">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-731">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-732">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-732">
         - PdfFile</span></span><br><span data-ttu-id="708a1-733">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-733">
         - Selection</span></span><br><span data-ttu-id="708a1-734">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-734">
         - Settings</span></span><br><span data-ttu-id="708a1-735">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-735">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-736">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-736">Office on Windows</span></span><br><span data-ttu-id="708a1-737">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-737">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-738">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-738">- Content</span></span><br><span data-ttu-id="708a1-739">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-739">
         - TaskPane</span></span><br><span data-ttu-id="708a1-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-742">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-742">- ActiveView</span></span><br><span data-ttu-id="708a1-743">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-743">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-744">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-744">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-745">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-745">
         - File</span></span><br><span data-ttu-id="708a1-746">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-746">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-747">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-747">
         - PdfFile</span></span><br><span data-ttu-id="708a1-748">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-748">
         - Selection</span></span><br><span data-ttu-id="708a1-749">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-749">
         - Settings</span></span><br><span data-ttu-id="708a1-750">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-750">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-751">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-751">Office 2019 on Windows</span></span><br><span data-ttu-id="708a1-752">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-752">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-753">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-753">- Content</span></span><br><span data-ttu-id="708a1-754">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-754">
         - TaskPane</span></span><br><span data-ttu-id="708a1-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-757">- ActiveView</span></span><br><span data-ttu-id="708a1-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-758">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-759">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-760">
         - File</span></span><br><span data-ttu-id="708a1-761">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-761">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-762">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-762">
         - PdfFile</span></span><br><span data-ttu-id="708a1-763">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-763">
         - Selection</span></span><br><span data-ttu-id="708a1-764">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-764">
         - Settings</span></span><br><span data-ttu-id="708a1-765">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-765">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-766">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-766">Office 2016 on Windows</span></span><br><span data-ttu-id="708a1-767">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-767">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-768">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-768">- Content</span></span><br><span data-ttu-id="708a1-769">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-769">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="708a1-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="708a1-771">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-771">- ActiveView</span></span><br><span data-ttu-id="708a1-772">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-772">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-773">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-773">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-774">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-774">
         - File</span></span><br><span data-ttu-id="708a1-775">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-775">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-776">
         - PdfFile</span></span><br><span data-ttu-id="708a1-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-777">
         - Selection</span></span><br><span data-ttu-id="708a1-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-778">
         - Settings</span></span><br><span data-ttu-id="708a1-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-780">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-780">Office 2013 on Windows</span></span><br><span data-ttu-id="708a1-781">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-782">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-782">- Content</span></span><br><span data-ttu-id="708a1-783">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-783">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="708a1-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="708a1-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="708a1-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-785">- ActiveView</span></span><br><span data-ttu-id="708a1-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-786">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-787">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-788">
         - File</span></span><br><span data-ttu-id="708a1-789">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-789">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-790">
         - PdfFile</span></span><br><span data-ttu-id="708a1-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-791">
         - Selection</span></span><br><span data-ttu-id="708a1-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-792">
         - Settings</span></span><br><span data-ttu-id="708a1-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-794">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="708a1-794">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="708a1-795">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-795">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-796">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-796">- Content</span></span><br><span data-ttu-id="708a1-797">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-797">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-799">- ActiveView</span></span><br><span data-ttu-id="708a1-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-800">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-801">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-802">
         - File</span></span><br><span data-ttu-id="708a1-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-803">
         - PdfFile</span></span><br><span data-ttu-id="708a1-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-804">
         - Selection</span></span><br><span data-ttu-id="708a1-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-805">
         - Settings</span></span><br><span data-ttu-id="708a1-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-806">
         - TextCoercion</span></span><br><span data-ttu-id="708a1-807">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-807">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-808">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-808">Office apps on Mac</span></span><br><span data-ttu-id="708a1-809">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="708a1-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="708a1-810">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-810">- Content</span></span><br><span data-ttu-id="708a1-811">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-811">
         - TaskPane</span></span><br><span data-ttu-id="708a1-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-814">- ActiveView</span></span><br><span data-ttu-id="708a1-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-815">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-816">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-817">
         - File</span></span><br><span data-ttu-id="708a1-818">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-818">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-819">
         - PdfFile</span></span><br><span data-ttu-id="708a1-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-820">
         - Selection</span></span><br><span data-ttu-id="708a1-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-821">
         - Settings</span></span><br><span data-ttu-id="708a1-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-823">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-823">Office 2019 for Mac</span></span><br><span data-ttu-id="708a1-824">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-824">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-825">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-825">- Content</span></span><br><span data-ttu-id="708a1-826">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-826">
         - TaskPane</span></span><br><span data-ttu-id="708a1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-829">- ActiveView</span></span><br><span data-ttu-id="708a1-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-830">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-831">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-832">
         - File</span></span><br><span data-ttu-id="708a1-833">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-833">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-834">
         - PdfFile</span></span><br><span data-ttu-id="708a1-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-835">
         - Selection</span></span><br><span data-ttu-id="708a1-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-836">
         - Settings</span></span><br><span data-ttu-id="708a1-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-838">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-838">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="708a1-839">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-840">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-840">- Content</span></span><br><span data-ttu-id="708a1-841">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-841">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="708a1-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="708a1-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="708a1-843">- ActiveView</span></span><br><span data-ttu-id="708a1-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="708a1-844">
         - CompressedFile</span></span><br><span data-ttu-id="708a1-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-845">
         - DocumentEvents</span></span><br><span data-ttu-id="708a1-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="708a1-846">
         - File</span></span><br><span data-ttu-id="708a1-847">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-847">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="708a1-848">
         - PdfFile</span></span><br><span data-ttu-id="708a1-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-849">
         - Selection</span></span><br><span data-ttu-id="708a1-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-850">
         - Settings</span></span><br><span data-ttu-id="708a1-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-851">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="708a1-852">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="708a1-852">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="708a1-853">OneNote</span><span class="sxs-lookup"><span data-stu-id="708a1-853">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="708a1-854">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="708a1-854">Platform</span></span></th>
    <th><span data-ttu-id="708a1-855">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-855">Extension points</span></span></th>
    <th><span data-ttu-id="708a1-856">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-856">API requirement sets</span></span></th>
    <th><span data-ttu-id="708a1-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-858">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="708a1-858">Office on the web</span></span></td>
    <td> <span data-ttu-id="708a1-859">- Contenu</span><span class="sxs-lookup"><span data-stu-id="708a1-859">- Content</span></span><br><span data-ttu-id="708a1-860">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-860">
         - TaskPane</span></span><br><span data-ttu-id="708a1-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="708a1-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="708a1-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="708a1-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-864">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="708a1-864">- DocumentEvents</span></span><br><span data-ttu-id="708a1-865">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-865">
         - HtmlCoercion</span></span><br><span data-ttu-id="708a1-866">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-866">
         - ImageCoercion</span></span><br><span data-ttu-id="708a1-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="708a1-867">
         - Settings</span></span><br><span data-ttu-id="708a1-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="708a1-869">Projet</span><span class="sxs-lookup"><span data-stu-id="708a1-869">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="708a1-870">Plateforme</span><span class="sxs-lookup"><span data-stu-id="708a1-870">Platform</span></span></th>
    <th><span data-ttu-id="708a1-871">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="708a1-871">Extension points</span></span></th>
    <th><span data-ttu-id="708a1-872">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="708a1-872">API requirement sets</span></span></th>
    <th><span data-ttu-id="708a1-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="708a1-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-874">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-874">Office 2019 on Windows</span></span><br><span data-ttu-id="708a1-875">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-875">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-876">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-876">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-878">- Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-878">- Selection</span></span><br><span data-ttu-id="708a1-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-879">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-880">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-880">Office 2016 on Windows</span></span><br><span data-ttu-id="708a1-881">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-881">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-882">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-882">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-884">- Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-884">- Selection</span></span><br><span data-ttu-id="708a1-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-885">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="708a1-886">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="708a1-886">Office 2013 on Windows</span></span><br><span data-ttu-id="708a1-887">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="708a1-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="708a1-888">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="708a1-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="708a1-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="708a1-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="708a1-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="708a1-890">- Selection</span></span><br><span data-ttu-id="708a1-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="708a1-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="708a1-892">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="708a1-892">See also</span></span>

- [<span data-ttu-id="708a1-893">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="708a1-893">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="708a1-894">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="708a1-894">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="708a1-895">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="708a1-895">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="708a1-896">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="708a1-896">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="708a1-897">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="708a1-897">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="708a1-898">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="708a1-898">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="708a1-899">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="708a1-899">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="708a1-900">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="708a1-900">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="708a1-901">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="708a1-901">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="708a1-902">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="708a1-902">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="708a1-903">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="708a1-903">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
