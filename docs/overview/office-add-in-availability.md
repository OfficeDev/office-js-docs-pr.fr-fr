---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: 1e368fe21a1bcdb2a7f44c88ce8e881605fa96f2
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395651"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="694bc-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="694bc-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="694bc-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="694bc-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="694bc-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="694bc-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="694bc-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="694bc-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="694bc-108">Excel</span><span class="sxs-lookup"><span data-stu-id="694bc-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="694bc-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="694bc-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="694bc-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="694bc-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="694bc-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="694bc-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-114">- TaskPane</span></span><br><span data-ttu-id="694bc-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-115">
        - Content</span></span><br><span data-ttu-id="694bc-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-116">
        - Custom Functions</span></span><br><span data-ttu-id="694bc-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="694bc-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="694bc-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="694bc-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="694bc-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-128">
        - BindingEvents</span></span><br><span data-ttu-id="694bc-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-129">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-130">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-131">
        - File</span></span><br><span data-ttu-id="694bc-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-132">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-134">
        - Selection</span></span><br><span data-ttu-id="694bc-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-135">
        - Settings</span></span><br><span data-ttu-id="694bc-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-136">
        - TableBindings</span></span><br><span data-ttu-id="694bc-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-137">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-138">
        - TextBindings</span></span><br><span data-ttu-id="694bc-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-140">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-140">Office on Windows</span></span><br><span data-ttu-id="694bc-141">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-142">- TaskPane</span></span><br><span data-ttu-id="694bc-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-143">
        - Content</span></span><br><span data-ttu-id="694bc-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-144">
        - Custom Functions</span></span><br><span data-ttu-id="694bc-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="694bc-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="694bc-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="694bc-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="694bc-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="694bc-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-158">
        - BindingEvents</span></span><br><span data-ttu-id="694bc-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-159">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-160">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-161">
        - File</span></span><br><span data-ttu-id="694bc-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-162">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-164">
        - Selection</span></span><br><span data-ttu-id="694bc-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-165">
        - Settings</span></span><br><span data-ttu-id="694bc-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-166">
        - TableBindings</span></span><br><span data-ttu-id="694bc-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-167">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-168">
        - TextBindings</span></span><br><span data-ttu-id="694bc-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-170">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-170">Office 2019 on Windows</span></span><br><span data-ttu-id="694bc-171">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="694bc-172">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-172">- TaskPane</span></span><br><span data-ttu-id="694bc-173">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-173">
        - Content</span></span><br><span data-ttu-id="694bc-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="694bc-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-185">- BindingEvents</span></span><br><span data-ttu-id="694bc-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-186">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-187">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-188">
        - File</span></span><br><span data-ttu-id="694bc-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-189">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-191">
        - Selection</span></span><br><span data-ttu-id="694bc-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-192">
        - Settings</span></span><br><span data-ttu-id="694bc-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-193">
        - TableBindings</span></span><br><span data-ttu-id="694bc-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-194">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-195">
        - TextBindings</span></span><br><span data-ttu-id="694bc-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-197">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-197">Office 2016 on Windows</span></span><br><span data-ttu-id="694bc-198">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="694bc-199">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-199">- TaskPane</span></span><br><span data-ttu-id="694bc-200">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-200">
        - Content</span></span></td>
    <td><span data-ttu-id="694bc-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="694bc-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-204">- BindingEvents</span></span><br><span data-ttu-id="694bc-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-205">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-206">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-207">
        - File</span></span><br><span data-ttu-id="694bc-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-208">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-210">
        - Selection</span></span><br><span data-ttu-id="694bc-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-211">
        - Settings</span></span><br><span data-ttu-id="694bc-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-212">
        - TableBindings</span></span><br><span data-ttu-id="694bc-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-213">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-214">
        - TextBindings</span></span><br><span data-ttu-id="694bc-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-216">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-216">Office 2013 on Windows</span></span><br><span data-ttu-id="694bc-217">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="694bc-218">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-218">
        - TaskPane</span></span><br><span data-ttu-id="694bc-219">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="694bc-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="694bc-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="694bc-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-222">
        - BindingEvents</span></span><br><span data-ttu-id="694bc-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-223">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-224">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-225">
        - File</span></span><br><span data-ttu-id="694bc-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-226">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-228">
        - Selection</span></span><br><span data-ttu-id="694bc-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-229">
        - Settings</span></span><br><span data-ttu-id="694bc-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-230">
        - TableBindings</span></span><br><span data-ttu-id="694bc-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-231">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-232">
        - TextBindings</span></span><br><span data-ttu-id="694bc-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-234">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="694bc-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="694bc-235">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="694bc-236">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-236">- TaskPane</span></span><br><span data-ttu-id="694bc-237">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-237">
        - Content</span></span><br><span data-ttu-id="694bc-238">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-238">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="694bc-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="694bc-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="694bc-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-250">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-250">- BindingEvents</span></span><br><span data-ttu-id="694bc-251">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-251">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-252">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-252">
        - File</span></span><br><span data-ttu-id="694bc-253">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-253">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-254">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-254">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-255">
        - Selection</span></span><br><span data-ttu-id="694bc-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-256">
        - Settings</span></span><br><span data-ttu-id="694bc-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-257">
        - TableBindings</span></span><br><span data-ttu-id="694bc-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-258">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-259">
        - TextBindings</span></span><br><span data-ttu-id="694bc-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-260">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-261">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-261">Office apps on Mac</span></span><br><span data-ttu-id="694bc-262">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-262">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="694bc-263">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-263">- TaskPane</span></span><br><span data-ttu-id="694bc-264">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-264">
        - Content</span></span><br><span data-ttu-id="694bc-265">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-265">
        - Custom Functions</span></span><br><span data-ttu-id="694bc-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="694bc-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="694bc-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="694bc-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="694bc-279">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-279">- BindingEvents</span></span><br><span data-ttu-id="694bc-280">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-280">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-281">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-281">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-282">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-282">
        - File</span></span><br><span data-ttu-id="694bc-283">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-283">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-284">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-284">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-285">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-285">
        - PdfFile</span></span><br><span data-ttu-id="694bc-286">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-286">
        - Selection</span></span><br><span data-ttu-id="694bc-287">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-287">
        - Settings</span></span><br><span data-ttu-id="694bc-288">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-288">
        - TableBindings</span></span><br><span data-ttu-id="694bc-289">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-289">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-290">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-290">
        - TextBindings</span></span><br><span data-ttu-id="694bc-291">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-291">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-292">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-292">Office 2019 for Mac</span></span><br><span data-ttu-id="694bc-293">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-293">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="694bc-294">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-294">- TaskPane</span></span><br><span data-ttu-id="694bc-295">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-295">
        - Content</span></span><br><span data-ttu-id="694bc-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="694bc-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="694bc-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="694bc-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="694bc-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="694bc-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="694bc-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="694bc-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="694bc-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="694bc-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-307">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-307">- BindingEvents</span></span><br><span data-ttu-id="694bc-308">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-308">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-309">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-309">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-310">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-310">
        - File</span></span><br><span data-ttu-id="694bc-311">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-311">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-312">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-312">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-313">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-313">
        - PdfFile</span></span><br><span data-ttu-id="694bc-314">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-314">
        - Selection</span></span><br><span data-ttu-id="694bc-315">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-315">
        - Settings</span></span><br><span data-ttu-id="694bc-316">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-316">
        - TableBindings</span></span><br><span data-ttu-id="694bc-317">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-317">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-318">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-318">
        - TextBindings</span></span><br><span data-ttu-id="694bc-319">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-319">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-320">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-320">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="694bc-321">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-321">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="694bc-322">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-322">- TaskPane</span></span><br><span data-ttu-id="694bc-323">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-323">
        - Content</span></span></td>
    <td><span data-ttu-id="694bc-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="694bc-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="694bc-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="694bc-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-327">- BindingEvents</span></span><br><span data-ttu-id="694bc-328">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-328">
        - CompressedFile</span></span><br><span data-ttu-id="694bc-329">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-329">
        - DocumentEvents</span></span><br><span data-ttu-id="694bc-330">
        - File</span><span class="sxs-lookup"><span data-stu-id="694bc-330">
        - File</span></span><br><span data-ttu-id="694bc-331">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-331">
        - MatrixBindings</span></span><br><span data-ttu-id="694bc-332">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-332">
        - MatrixCoercion</span></span><br><span data-ttu-id="694bc-333">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-333">
        - PdfFile</span></span><br><span data-ttu-id="694bc-334">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-334">
        - Selection</span></span><br><span data-ttu-id="694bc-335">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-335">
        - Settings</span></span><br><span data-ttu-id="694bc-336">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-336">
        - TableBindings</span></span><br><span data-ttu-id="694bc-337">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-337">
        - TableCoercion</span></span><br><span data-ttu-id="694bc-338">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-338">
        - TextBindings</span></span><br><span data-ttu-id="694bc-339">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-339">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="694bc-340">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="694bc-340">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="694bc-341">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-341">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="694bc-342">Plateforme</span><span class="sxs-lookup"><span data-stu-id="694bc-342">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="694bc-343">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-343">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="694bc-344">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-344">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="694bc-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-346">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-346">Office on the web</span></span></td>
    <td><span data-ttu-id="694bc-347">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-347">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="694bc-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-349">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-349">Office on Windows</span></span><br><span data-ttu-id="694bc-350">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-350">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="694bc-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="694bc-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-353">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-353">Office for Mac</span></span><br><span data-ttu-id="694bc-354">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-354">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="694bc-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="694bc-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="694bc-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="694bc-357">Outlook</span><span class="sxs-lookup"><span data-stu-id="694bc-357">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="694bc-358">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="694bc-358">Platform</span></span></th>
    <th><span data-ttu-id="694bc-359">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-359">Extension points</span></span></th>
    <th><span data-ttu-id="694bc-360">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-360">API requirement sets</span></span></th>
    <th><span data-ttu-id="694bc-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-362">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-362">Office on the web</span></span><br><span data-ttu-id="694bc-363">(moderne)</span><span class="sxs-lookup"><span data-stu-id="694bc-363">Modern</span></span></td>
    <td> <span data-ttu-id="694bc-364">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-364">- Mail Read</span></span><br><span data-ttu-id="694bc-365">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-365">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="694bc-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="694bc-374">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-375">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-375">Office on the web</span></span><br><span data-ttu-id="694bc-376">(classique)</span><span class="sxs-lookup"><span data-stu-id="694bc-376">Classic.</span></span></td>
    <td> <span data-ttu-id="694bc-377">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-377">- Mail Read</span></span><br><span data-ttu-id="694bc-378">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-378">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="694bc-386">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-387">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-387">Office on Windows</span></span><br><span data-ttu-id="694bc-388">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-389">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-389">- Mail Read</span></span><br><span data-ttu-id="694bc-390">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-390">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="694bc-392">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="694bc-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="694bc-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="694bc-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="694bc-400">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-400">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-401">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-401">Office 2019 on Windows</span></span><br><span data-ttu-id="694bc-402">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-402">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-403">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-403">- Mail Read</span></span><br><span data-ttu-id="694bc-404">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-404">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="694bc-406">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="694bc-406">
      - Modules</span></span></td>
    <td> <span data-ttu-id="694bc-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="694bc-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="694bc-414">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-415">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-415">Office 2016 on Windows</span></span><br><span data-ttu-id="694bc-416">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-416">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-417">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-417">- Mail Read</span></span><br><span data-ttu-id="694bc-418">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-418">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="694bc-420">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="694bc-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="694bc-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="694bc-425">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-426">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-426">Office 2013 on Windows</span></span><br><span data-ttu-id="694bc-427">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-427">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-428">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-428">- Mail Read</span></span><br><span data-ttu-id="694bc-429">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-429">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="694bc-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="694bc-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="694bc-434">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-434">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-435">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="694bc-435">Office apps on iOS</span></span><br><span data-ttu-id="694bc-436">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-436">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-437">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-437">- Mail Read</span></span><br><span data-ttu-id="694bc-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="694bc-444">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-445">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-445">Office apps on Mac</span></span><br><span data-ttu-id="694bc-446">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-446">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-447">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-447">- Mail Read</span></span><br><span data-ttu-id="694bc-448">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-448">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="694bc-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="694bc-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="694bc-457">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-458">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-458">Office 2019 for Mac</span></span><br><span data-ttu-id="694bc-459">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-460">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-460">- Mail Read</span></span><br><span data-ttu-id="694bc-461">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-461">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="694bc-469">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-470">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-470">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="694bc-471">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-471">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-472">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-472">- Mail Read</span></span><br><span data-ttu-id="694bc-473">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="694bc-473">
      - Mail Compose</span></span><br><span data-ttu-id="694bc-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="694bc-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="694bc-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="694bc-481">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-482">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="694bc-482">Office apps on Android</span></span><br><span data-ttu-id="694bc-483">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-483">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-484">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="694bc-484">- Mail Read</span></span><br><span data-ttu-id="694bc-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="694bc-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="694bc-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="694bc-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="694bc-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="694bc-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="694bc-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="694bc-491">Non disponible</span><span class="sxs-lookup"><span data-stu-id="694bc-491">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="694bc-492">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="694bc-492">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="694bc-493">Word</span><span class="sxs-lookup"><span data-stu-id="694bc-493">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="694bc-494">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="694bc-494">Platform</span></span></th>
    <th><span data-ttu-id="694bc-495">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-495">Extension points</span></span></th>
    <th><span data-ttu-id="694bc-496">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-496">API requirement sets</span></span></th>
    <th><span data-ttu-id="694bc-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-498">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-498">Office on the web</span></span></td>
    <td> <span data-ttu-id="694bc-499">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-499">- TaskPane</span></span><br><span data-ttu-id="694bc-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="694bc-507">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-507">- BindingEvents</span></span><br><span data-ttu-id="694bc-508">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-508">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-509">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-509">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-510">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-510">
         - File</span></span><br><span data-ttu-id="694bc-511">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-511">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-512">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-512">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-513">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-513">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-514">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-514">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-515">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-515">
         - PdfFile</span></span><br><span data-ttu-id="694bc-516">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-516">
         - Selection</span></span><br><span data-ttu-id="694bc-517">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-517">
         - Settings</span></span><br><span data-ttu-id="694bc-518">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-518">
         - TableBindings</span></span><br><span data-ttu-id="694bc-519">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-519">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-520">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-520">
         - TextBindings</span></span><br><span data-ttu-id="694bc-521">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-521">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-522">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-522">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-523">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-523">Office on Windows</span></span><br><span data-ttu-id="694bc-524">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-524">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-525">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-525">- TaskPane</span></span><br><span data-ttu-id="694bc-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="694bc-533">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-533">- BindingEvents</span></span><br><span data-ttu-id="694bc-534">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-534">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-536">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-537">
         - File</span></span><br><span data-ttu-id="694bc-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-539">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-542">
         - PdfFile</span></span><br><span data-ttu-id="694bc-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-543">
         - Selection</span></span><br><span data-ttu-id="694bc-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-544">
         - Settings</span></span><br><span data-ttu-id="694bc-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-545">
         - TableBindings</span></span><br><span data-ttu-id="694bc-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-546">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-547">
         - TextBindings</span></span><br><span data-ttu-id="694bc-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-548">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-549">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-550">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-550">Office 2019 on Windows</span></span><br><span data-ttu-id="694bc-551">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-551">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-552">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="694bc-552">- TaskPane</span></span><br><span data-ttu-id="694bc-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-559">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-559">- BindingEvents</span></span><br><span data-ttu-id="694bc-560">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-560">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-561">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-561">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-562">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-562">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-563">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-563">
         - File</span></span><br><span data-ttu-id="694bc-564">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-564">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-565">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-565">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-566">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-566">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-567">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-567">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-568">
         - PdfFile</span></span><br><span data-ttu-id="694bc-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-569">
         - Selection</span></span><br><span data-ttu-id="694bc-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-570">
         - Settings</span></span><br><span data-ttu-id="694bc-571">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-571">
         - TableBindings</span></span><br><span data-ttu-id="694bc-572">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-572">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-573">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-573">
         - TextBindings</span></span><br><span data-ttu-id="694bc-574">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-574">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-575">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-575">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-576">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-576">Office 2016 on Windows</span></span><br><span data-ttu-id="694bc-577">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-577">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-578">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-578">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="694bc-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-582">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-582">- BindingEvents</span></span><br><span data-ttu-id="694bc-583">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-583">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-584">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-584">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-585">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-586">
         - File</span></span><br><span data-ttu-id="694bc-587">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-587">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-588">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-588">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-589">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-589">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-590">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-590">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-591">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-591">
         - PdfFile</span></span><br><span data-ttu-id="694bc-592">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-592">
         - Selection</span></span><br><span data-ttu-id="694bc-593">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-593">
         - Settings</span></span><br><span data-ttu-id="694bc-594">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-594">
         - TableBindings</span></span><br><span data-ttu-id="694bc-595">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-595">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-596">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-596">
         - TextBindings</span></span><br><span data-ttu-id="694bc-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-597">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-598">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-598">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-599">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-599">Office 2013 on Windows</span></span><br><span data-ttu-id="694bc-600">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-600">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-601">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-601">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="694bc-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="694bc-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-604">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-604">- BindingEvents</span></span><br><span data-ttu-id="694bc-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-605">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-606">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-606">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-607">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-608">
         - File</span></span><br><span data-ttu-id="694bc-609">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-609">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-610">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-610">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-611">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-611">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-612">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-612">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-613">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-613">
         - PdfFile</span></span><br><span data-ttu-id="694bc-614">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-614">
         - Selection</span></span><br><span data-ttu-id="694bc-615">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-615">
         - Settings</span></span><br><span data-ttu-id="694bc-616">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-616">
         - TableBindings</span></span><br><span data-ttu-id="694bc-617">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-617">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-618">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-618">
         - TextBindings</span></span><br><span data-ttu-id="694bc-619">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-619">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-620">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-620">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-621">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="694bc-621">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="694bc-622">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-622">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-623">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-623">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="694bc-629">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-629">- BindingEvents</span></span><br><span data-ttu-id="694bc-630">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-630">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-631">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-631">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-632">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-633">
         - File</span></span><br><span data-ttu-id="694bc-634">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-634">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-635">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-635">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-636">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-636">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-637">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-637">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-638">
         - PdfFile</span></span><br><span data-ttu-id="694bc-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-639">
         - Selection</span></span><br><span data-ttu-id="694bc-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-640">
         - Settings</span></span><br><span data-ttu-id="694bc-641">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-641">
         - TableBindings</span></span><br><span data-ttu-id="694bc-642">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-642">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-643">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-643">
         - TextBindings</span></span><br><span data-ttu-id="694bc-644">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-644">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-645">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-645">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-646">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-646">Office apps on Mac</span></span><br><span data-ttu-id="694bc-647">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-647">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-648">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-648">- TaskPane</span></span><br><span data-ttu-id="694bc-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="694bc-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-656">- BindingEvents</span></span><br><span data-ttu-id="694bc-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-657">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-659">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-660">
         - File</span></span><br><span data-ttu-id="694bc-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-662">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-665">
         - PdfFile</span></span><br><span data-ttu-id="694bc-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-666">
         - Selection</span></span><br><span data-ttu-id="694bc-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-667">
         - Settings</span></span><br><span data-ttu-id="694bc-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-668">
         - TableBindings</span></span><br><span data-ttu-id="694bc-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-669">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-670">
         - TextBindings</span></span><br><span data-ttu-id="694bc-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-671">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-673">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-673">Office 2019 for Mac</span></span><br><span data-ttu-id="694bc-674">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-674">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-675">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="694bc-675">- TaskPane</span></span><br><span data-ttu-id="694bc-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="694bc-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="694bc-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="694bc-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="694bc-682">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-682">- BindingEvents</span></span><br><span data-ttu-id="694bc-683">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-683">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-684">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-684">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-685">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-685">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-686">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-686">
         - File</span></span><br><span data-ttu-id="694bc-687">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-687">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-688">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-688">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-689">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-689">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-690">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-690">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-691">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-691">
         - PdfFile</span></span><br><span data-ttu-id="694bc-692">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-692">
         - Selection</span></span><br><span data-ttu-id="694bc-693">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-693">
         - Settings</span></span><br><span data-ttu-id="694bc-694">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-694">
         - TableBindings</span></span><br><span data-ttu-id="694bc-695">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-695">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-696">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-696">
         - TextBindings</span></span><br><span data-ttu-id="694bc-697">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-697">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-698">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-698">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-699">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-699">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="694bc-700">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-700">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-701">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-701">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="694bc-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="694bc-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="694bc-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-705">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-705">- BindingEvents</span></span><br><span data-ttu-id="694bc-706">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-706">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-707">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="694bc-707">
         - CustomXmlParts</span></span><br><span data-ttu-id="694bc-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-708">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-709">
         - File</span></span><br><span data-ttu-id="694bc-710">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-710">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-711">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-711">
         - MatrixBindings</span></span><br><span data-ttu-id="694bc-712">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-712">
         - MatrixCoercion</span></span><br><span data-ttu-id="694bc-713">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-713">
         - OoxmlCoercion</span></span><br><span data-ttu-id="694bc-714">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-714">
         - PdfFile</span></span><br><span data-ttu-id="694bc-715">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-715">
         - Selection</span></span><br><span data-ttu-id="694bc-716">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-716">
         - Settings</span></span><br><span data-ttu-id="694bc-717">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-717">
         - TableBindings</span></span><br><span data-ttu-id="694bc-718">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-718">
         - TableCoercion</span></span><br><span data-ttu-id="694bc-719">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="694bc-719">
         - TextBindings</span></span><br><span data-ttu-id="694bc-720">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-720">
         - TextCoercion</span></span><br><span data-ttu-id="694bc-721">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="694bc-721">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="694bc-722">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="694bc-722">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="694bc-723">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="694bc-723">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="694bc-724">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="694bc-724">Platform</span></span></th>
    <th><span data-ttu-id="694bc-725">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-725">Extension points</span></span></th>
    <th><span data-ttu-id="694bc-726">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-726">API requirement sets</span></span></th>
    <th><span data-ttu-id="694bc-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-728">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-728">Office on the web</span></span></td>
    <td> <span data-ttu-id="694bc-729">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-729">- Content</span></span><br><span data-ttu-id="694bc-730">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-730">
         - TaskPane</span></span><br><span data-ttu-id="694bc-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="694bc-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="694bc-736">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-736">- ActiveView</span></span><br><span data-ttu-id="694bc-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-737">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-738">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-738">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-739">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-739">
         - File</span></span><br><span data-ttu-id="694bc-740">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-740">
         - PdfFile</span></span><br><span data-ttu-id="694bc-741">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-741">
         - Selection</span></span><br><span data-ttu-id="694bc-742">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-742">
         - Settings</span></span><br><span data-ttu-id="694bc-743">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-743">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-744">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-744">Office on Windows</span></span><br><span data-ttu-id="694bc-745">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-745">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-746">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-746">- Content</span></span><br><span data-ttu-id="694bc-747">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-747">
         - TaskPane</span></span><br><span data-ttu-id="694bc-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="694bc-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="694bc-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-753">- ActiveView</span></span><br><span data-ttu-id="694bc-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-754">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-755">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-756">
         - File</span></span><br><span data-ttu-id="694bc-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-757">
         - PdfFile</span></span><br><span data-ttu-id="694bc-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-758">
         - Selection</span></span><br><span data-ttu-id="694bc-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-759">
         - Settings</span></span><br><span data-ttu-id="694bc-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-761">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-761">Office 2019 on Windows</span></span><br><span data-ttu-id="694bc-762">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-763">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-763">- Content</span></span><br><span data-ttu-id="694bc-764">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-764">
         - TaskPane</span></span><br><span data-ttu-id="694bc-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-768">- ActiveView</span></span><br><span data-ttu-id="694bc-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-769">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-770">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-771">
         - File</span></span><br><span data-ttu-id="694bc-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-772">
         - PdfFile</span></span><br><span data-ttu-id="694bc-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-773">
         - Selection</span></span><br><span data-ttu-id="694bc-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-774">
         - Settings</span></span><br><span data-ttu-id="694bc-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-776">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-776">Office 2016 on Windows</span></span><br><span data-ttu-id="694bc-777">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-778">- Content</span></span><br><span data-ttu-id="694bc-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="694bc-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="694bc-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-782">- ActiveView</span></span><br><span data-ttu-id="694bc-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-783">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-784">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-785">
         - File</span></span><br><span data-ttu-id="694bc-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-786">
         - PdfFile</span></span><br><span data-ttu-id="694bc-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-787">
         - Selection</span></span><br><span data-ttu-id="694bc-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-788">
         - Settings</span></span><br><span data-ttu-id="694bc-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-790">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-790">Office 2013 on Windows</span></span><br><span data-ttu-id="694bc-791">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-792">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-792">- Content</span></span><br><span data-ttu-id="694bc-793">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="694bc-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="694bc-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="694bc-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-796">- ActiveView</span></span><br><span data-ttu-id="694bc-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-797">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-798">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-799">
         - File</span></span><br><span data-ttu-id="694bc-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-800">
         - PdfFile</span></span><br><span data-ttu-id="694bc-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-801">
         - Selection</span></span><br><span data-ttu-id="694bc-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-802">
         - Settings</span></span><br><span data-ttu-id="694bc-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-804">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="694bc-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="694bc-805">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-806">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-806">- Content</span></span><br><span data-ttu-id="694bc-807">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="694bc-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-811">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-811">- ActiveView</span></span><br><span data-ttu-id="694bc-812">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-812">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-813">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-813">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-814">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-814">
         - File</span></span><br><span data-ttu-id="694bc-815">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-815">
         - PdfFile</span></span><br><span data-ttu-id="694bc-816">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-816">
         - Selection</span></span><br><span data-ttu-id="694bc-817">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-817">
         - Settings</span></span><br><span data-ttu-id="694bc-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-818">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-819">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-819">Office apps on Mac</span></span><br><span data-ttu-id="694bc-820">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="694bc-820">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="694bc-821">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-821">- Content</span></span><br><span data-ttu-id="694bc-822">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-822">
         - TaskPane</span></span><br><span data-ttu-id="694bc-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="694bc-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="694bc-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="694bc-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="694bc-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-828">- ActiveView</span></span><br><span data-ttu-id="694bc-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-829">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-830">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-831">
         - File</span></span><br><span data-ttu-id="694bc-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-832">
         - PdfFile</span></span><br><span data-ttu-id="694bc-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-833">
         - Selection</span></span><br><span data-ttu-id="694bc-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-834">
         - Settings</span></span><br><span data-ttu-id="694bc-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-836">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-836">Office 2019 for Mac</span></span><br><span data-ttu-id="694bc-837">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-837">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-838">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-838">- Content</span></span><br><span data-ttu-id="694bc-839">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-839">
         - TaskPane</span></span><br><span data-ttu-id="694bc-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-843">- ActiveView</span></span><br><span data-ttu-id="694bc-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-844">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-845">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-846">
         - File</span></span><br><span data-ttu-id="694bc-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-847">
         - PdfFile</span></span><br><span data-ttu-id="694bc-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-848">
         - Selection</span></span><br><span data-ttu-id="694bc-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-849">
         - Settings</span></span><br><span data-ttu-id="694bc-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-851">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-851">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="694bc-852">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-852">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-853">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-853">- Content</span></span><br><span data-ttu-id="694bc-854">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-854">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="694bc-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="694bc-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-857">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="694bc-857">- ActiveView</span></span><br><span data-ttu-id="694bc-858">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="694bc-858">
         - CompressedFile</span></span><br><span data-ttu-id="694bc-859">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-859">
         - DocumentEvents</span></span><br><span data-ttu-id="694bc-860">
         - File</span><span class="sxs-lookup"><span data-stu-id="694bc-860">
         - File</span></span><br><span data-ttu-id="694bc-861">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="694bc-861">
         - PdfFile</span></span><br><span data-ttu-id="694bc-862">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-862">
         - Selection</span></span><br><span data-ttu-id="694bc-863">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-863">
         - Settings</span></span><br><span data-ttu-id="694bc-864">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-864">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="694bc-865">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="694bc-865">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="694bc-866">OneNote</span><span class="sxs-lookup"><span data-stu-id="694bc-866">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="694bc-867">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="694bc-867">Platform</span></span></th>
    <th><span data-ttu-id="694bc-868">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-868">Extension points</span></span></th>
    <th><span data-ttu-id="694bc-869">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-869">API requirement sets</span></span></th>
    <th><span data-ttu-id="694bc-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-871">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="694bc-871">Office on the web</span></span></td>
    <td> <span data-ttu-id="694bc-872">- Contenu</span><span class="sxs-lookup"><span data-stu-id="694bc-872">- Content</span></span><br><span data-ttu-id="694bc-873">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-873">
         - TaskPane</span></span><br><span data-ttu-id="694bc-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="694bc-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="694bc-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="694bc-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="694bc-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-878">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="694bc-878">- DocumentEvents</span></span><br><span data-ttu-id="694bc-879">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-879">
         - HtmlCoercion</span></span><br><span data-ttu-id="694bc-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="694bc-880">
         - Settings</span></span><br><span data-ttu-id="694bc-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="694bc-882">Projet</span><span class="sxs-lookup"><span data-stu-id="694bc-882">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="694bc-883">Plateforme</span><span class="sxs-lookup"><span data-stu-id="694bc-883">Platform</span></span></th>
    <th><span data-ttu-id="694bc-884">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="694bc-884">Extension points</span></span></th>
    <th><span data-ttu-id="694bc-885">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="694bc-885">API requirement sets</span></span></th>
    <th><span data-ttu-id="694bc-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="694bc-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-887">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-887">Office 2019 on Windows</span></span><br><span data-ttu-id="694bc-888">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-888">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-889">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-889">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-891">- Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-891">- Selection</span></span><br><span data-ttu-id="694bc-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-892">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-893">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-893">Office 2016 on Windows</span></span><br><span data-ttu-id="694bc-894">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-894">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-895">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-895">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-897">- Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-897">- Selection</span></span><br><span data-ttu-id="694bc-898">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-898">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="694bc-899">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="694bc-899">Office 2013 on Windows</span></span><br><span data-ttu-id="694bc-900">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="694bc-900">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="694bc-901">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="694bc-901">- TaskPane</span></span></td>
    <td> <span data-ttu-id="694bc-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="694bc-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="694bc-903">- Selection</span><span class="sxs-lookup"><span data-stu-id="694bc-903">- Selection</span></span><br><span data-ttu-id="694bc-904">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="694bc-904">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="694bc-905">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="694bc-905">See also</span></span>

- [<span data-ttu-id="694bc-906">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="694bc-906">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="694bc-907">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="694bc-907">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="694bc-908">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="694bc-908">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="694bc-909">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="694bc-909">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="694bc-910">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="694bc-910">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="694bc-911">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="694bc-911">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="694bc-912">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="694bc-912">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="694bc-913">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="694bc-913">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="694bc-914">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="694bc-914">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="694bc-915">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="694bc-915">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="694bc-916">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="694bc-916">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
