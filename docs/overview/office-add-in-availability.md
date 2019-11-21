---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: ecb906e595c08b973b5146416a5317d59547ed39
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757484"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c6a0c-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="c6a0c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c6a0c-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="c6a0c-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c6a0c-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="c6a0c-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c6a0c-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="c6a0c-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c6a0c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c6a0c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c6a0c-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c6a0c-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c6a0c-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c6a0c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c6a0c-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-114">- TaskPane</span></span><br><span data-ttu-id="c6a0c-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-115">
        - Content</span></span><br><span data-ttu-id="c6a0c-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-116">
        - Custom Functions</span></span><br><span data-ttu-id="c6a0c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="c6a0c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c6a0c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c6a0c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c6a0c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c6a0c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-130">
        - BindingEvents</span></span><br><span data-ttu-id="c6a0c-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-131">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-132">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-133">
        - File</span></span><br><span data-ttu-id="c6a0c-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-134">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-136">
        - Selection</span></span><br><span data-ttu-id="c6a0c-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-137">
        - Settings</span></span><br><span data-ttu-id="c6a0c-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-138">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-139">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-140">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-142">Office on Windows</span></span><br><span data-ttu-id="c6a0c-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-144">- TaskPane</span></span><br><span data-ttu-id="c6a0c-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-145">
        - Content</span></span><br><span data-ttu-id="c6a0c-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-146">
        - Custom Functions</span></span><br><span data-ttu-id="c6a0c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="c6a0c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c6a0c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c6a0c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c6a0c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c6a0c-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-161">
        - BindingEvents</span></span><br><span data-ttu-id="c6a0c-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-162">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-163">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-164">
        - File</span></span><br><span data-ttu-id="c6a0c-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-165">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-167">
        - Selection</span></span><br><span data-ttu-id="c6a0c-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-168">
        - Settings</span></span><br><span data-ttu-id="c6a0c-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-169">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-170">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-171">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-173">Office 2019 on Windows</span></span><br><span data-ttu-id="c6a0c-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c6a0c-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-175">- TaskPane</span></span><br><span data-ttu-id="c6a0c-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-176">
        - Content</span></span><br><span data-ttu-id="c6a0c-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c6a0c-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-188">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-189">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-190">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-191">
        - File</span></span><br><span data-ttu-id="c6a0c-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-192">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-194">
        - Selection</span></span><br><span data-ttu-id="c6a0c-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-195">
        - Settings</span></span><br><span data-ttu-id="c6a0c-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-196">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-197">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-198">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-200">Office 2016 on Windows</span></span><br><span data-ttu-id="c6a0c-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c6a0c-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-202">- TaskPane</span></span><br><span data-ttu-id="c6a0c-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-203">
        - Content</span></span></td>
    <td><span data-ttu-id="c6a0c-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c6a0c-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-207">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-208">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-209">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-210">
        - File</span></span><br><span data-ttu-id="c6a0c-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-211">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-213">
        - Selection</span></span><br><span data-ttu-id="c6a0c-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-214">
        - Settings</span></span><br><span data-ttu-id="c6a0c-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-215">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-216">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-217">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-219">Office 2013 on Windows</span></span><br><span data-ttu-id="c6a0c-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c6a0c-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-221">
        - TaskPane</span></span><br><span data-ttu-id="c6a0c-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c6a0c-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c6a0c-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-225">
        - BindingEvents</span></span><br><span data-ttu-id="c6a0c-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-226">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-227">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-228">
        - File</span></span><br><span data-ttu-id="c6a0c-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-229">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-231">
        - Selection</span></span><br><span data-ttu-id="c6a0c-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-232">
        - Settings</span></span><br><span data-ttu-id="c6a0c-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-233">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-234">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-235">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="c6a0c-237">Office on iPad</span></span><br><span data-ttu-id="c6a0c-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c6a0c-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-239">- TaskPane</span></span><br><span data-ttu-id="c6a0c-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-240">
        - Content</span></span></td>
    <td><span data-ttu-id="c6a0c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c6a0c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c6a0c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-253">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-254">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-255">
        - File</span></span><br><span data-ttu-id="c6a0c-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-256">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-258">
        - Selection</span></span><br><span data-ttu-id="c6a0c-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-259">
        - Settings</span></span><br><span data-ttu-id="c6a0c-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-260">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-261">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-262">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-264">Office on Mac</span></span><br><span data-ttu-id="c6a0c-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c6a0c-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-266">- TaskPane</span></span><br><span data-ttu-id="c6a0c-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-267">
        - Content</span></span><br><span data-ttu-id="c6a0c-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-268">
        - Custom Functions</span></span><br><span data-ttu-id="c6a0c-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c6a0c-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c6a0c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c6a0c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c6a0c-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-283">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-284">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-285">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-286">
        - File</span></span><br><span data-ttu-id="c6a0c-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-287">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-289">
        - PdfFile</span></span><br><span data-ttu-id="c6a0c-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-290">
        - Selection</span></span><br><span data-ttu-id="c6a0c-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-291">
        - Settings</span></span><br><span data-ttu-id="c6a0c-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-292">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-293">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-294">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-296">Office 2019 on Mac</span></span><br><span data-ttu-id="c6a0c-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c6a0c-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-298">- TaskPane</span></span><br><span data-ttu-id="c6a0c-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-299">
        - Content</span></span><br><span data-ttu-id="c6a0c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c6a0c-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c6a0c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c6a0c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c6a0c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c6a0c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c6a0c-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-311">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-312">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-313">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-314">
        - File</span></span><br><span data-ttu-id="c6a0c-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-315">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-317">
        - PdfFile</span></span><br><span data-ttu-id="c6a0c-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-318">
        - Selection</span></span><br><span data-ttu-id="c6a0c-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-319">
        - Settings</span></span><br><span data-ttu-id="c6a0c-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-320">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-321">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-322">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-324">Office 2016 on Mac</span></span><br><span data-ttu-id="c6a0c-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c6a0c-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-326">- TaskPane</span></span><br><span data-ttu-id="c6a0c-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-327">
        - Content</span></span></td>
    <td><span data-ttu-id="c6a0c-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c6a0c-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c6a0c-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-331">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-332">
        - CompressedFile</span></span><br><span data-ttu-id="c6a0c-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-333">
        - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-334">
        - File</span></span><br><span data-ttu-id="c6a0c-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-335">
        - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-337">
        - PdfFile</span></span><br><span data-ttu-id="c6a0c-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-338">
        - Selection</span></span><br><span data-ttu-id="c6a0c-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-339">
        - Settings</span></span><br><span data-ttu-id="c6a0c-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-340">
        - TableBindings</span></span><br><span data-ttu-id="c6a0c-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-341">
        - TableCoercion</span></span><br><span data-ttu-id="c6a0c-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-342">
        - TextBindings</span></span><br><span data-ttu-id="c6a0c-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c6a0c-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="c6a0c-345">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c6a0c-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c6a0c-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c6a0c-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c6a0c-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-350">Office on the web</span></span></td>
    <td><span data-ttu-id="c6a0c-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c6a0c-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-353">Office on Windows</span></span><br><span data-ttu-id="c6a0c-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c6a0c-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c6a0c-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-357">Office for Mac</span></span><br><span data-ttu-id="c6a0c-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c6a0c-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c6a0c-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c6a0c-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c6a0c-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="c6a0c-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c6a0c-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-362">Platform</span></span></th>
    <th><span data-ttu-id="c6a0c-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-363">Extension points</span></span></th>
    <th><span data-ttu-id="c6a0c-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="c6a0c-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-366">Office on the web</span></span><br><span data-ttu-id="c6a0c-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-367">(modern)</span></span></td>
    <td> <span data-ttu-id="c6a0c-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-368">- Mail Read</span></span><br><span data-ttu-id="c6a0c-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-369">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c6a0c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c6a0c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c6a0c-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-380">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-380">Office on the web</span></span><br><span data-ttu-id="c6a0c-381">(classique)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-381">(classic)</span></span></td>
    <td> <span data-ttu-id="c6a0c-382">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-382">- Mail Read</span></span><br><span data-ttu-id="c6a0c-383">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-383">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c6a0c-391">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-392">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-392">Office on Windows</span></span><br><span data-ttu-id="c6a0c-393">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-394">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-394">- Mail Read</span></span><br><span data-ttu-id="c6a0c-395">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-395">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c6a0c-397">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="c6a0c-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c6a0c-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c6a0c-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c6a0c-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c6a0c-406">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-407">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-407">Office 2019 on Windows</span></span><br><span data-ttu-id="c6a0c-408">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-409">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-409">- Mail Read</span></span><br><span data-ttu-id="c6a0c-410">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-410">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c6a0c-412">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="c6a0c-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c6a0c-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c6a0c-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c6a0c-420">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-421">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-421">Office 2016 on Windows</span></span><br><span data-ttu-id="c6a0c-422">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-423">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-423">- Mail Read</span></span><br><span data-ttu-id="c6a0c-424">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-424">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c6a0c-426">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="c6a0c-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c6a0c-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c6a0c-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-432">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-432">Office 2013 on Windows</span></span><br><span data-ttu-id="c6a0c-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-434">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-434">- Mail Read</span></span><br><span data-ttu-id="c6a0c-435">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="c6a0c-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c6a0c-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c6a0c-440">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-441">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="c6a0c-441">Office on iOS</span></span><br><span data-ttu-id="c6a0c-442">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-443">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-443">- Mail Read</span></span><br><span data-ttu-id="c6a0c-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c6a0c-450">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-451">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-451">Office on Mac</span></span><br><span data-ttu-id="c6a0c-452">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-453">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-453">- Mail Read</span></span><br><span data-ttu-id="c6a0c-454">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-454">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c6a0c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c6a0c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c6a0c-464">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-465">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-465">Office 2019 on Mac</span></span><br><span data-ttu-id="c6a0c-466">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-467">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-467">- Mail Read</span></span><br><span data-ttu-id="c6a0c-468">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-468">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c6a0c-476">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-477">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-477">Office 2016 on Mac</span></span><br><span data-ttu-id="c6a0c-478">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-479">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-479">- Mail Read</span></span><br><span data-ttu-id="c6a0c-480">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-480">
      - Mail Compose</span></span><br><span data-ttu-id="c6a0c-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c6a0c-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c6a0c-488">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-489">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="c6a0c-489">Office on Android</span></span><br><span data-ttu-id="c6a0c-490">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-491">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c6a0c-491">- Mail Read</span></span><br><span data-ttu-id="c6a0c-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c6a0c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c6a0c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c6a0c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c6a0c-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c6a0c-498">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c6a0c-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c6a0c-499">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c6a0c-500">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c6a0c-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c6a0c-501">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6a0c-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c6a0c-502">Word</span><span class="sxs-lookup"><span data-stu-id="c6a0c-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c6a0c-503">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-503">Platform</span></span></th>
    <th><span data-ttu-id="c6a0c-504">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-504">Extension points</span></span></th>
    <th><span data-ttu-id="c6a0c-505">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="c6a0c-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-507">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="c6a0c-508">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-508">- TaskPane</span></span><br><span data-ttu-id="c6a0c-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-516">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-518">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-519">
         - File</span></span><br><span data-ttu-id="c6a0c-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-521">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-524">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-525">
         - Selection</span></span><br><span data-ttu-id="c6a0c-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-526">
         - Settings</span></span><br><span data-ttu-id="c6a0c-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-527">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-528">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-529">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-530">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-532">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-532">Office on Windows</span></span><br><span data-ttu-id="c6a0c-533">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-534">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-534">- TaskPane</span></span><br><span data-ttu-id="c6a0c-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-542">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-543">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-545">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-546">
         - File</span></span><br><span data-ttu-id="c6a0c-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-548">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-551">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-552">
         - Selection</span></span><br><span data-ttu-id="c6a0c-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-553">
         - Settings</span></span><br><span data-ttu-id="c6a0c-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-554">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-555">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-556">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-557">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-559">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-559">Office 2019 on Windows</span></span><br><span data-ttu-id="c6a0c-560">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-561">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="c6a0c-561">- TaskPane</span></span><br><span data-ttu-id="c6a0c-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-568">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-569">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-571">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-572">
         - File</span></span><br><span data-ttu-id="c6a0c-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-574">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-577">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-578">
         - Selection</span></span><br><span data-ttu-id="c6a0c-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-579">
         - Settings</span></span><br><span data-ttu-id="c6a0c-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-580">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-581">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-582">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-583">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-585">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-585">Office 2016 on Windows</span></span><br><span data-ttu-id="c6a0c-586">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-587">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c6a0c-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-591">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-592">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-595">
         - File</span></span><br><span data-ttu-id="c6a0c-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-600">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-601">
         - Selection</span></span><br><span data-ttu-id="c6a0c-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-602">
         - Settings</span></span><br><span data-ttu-id="c6a0c-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-603">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-604">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-605">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-606">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-608">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-608">Office 2013 on Windows</span></span><br><span data-ttu-id="c6a0c-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c6a0c-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-613">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-614">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-616">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-617">
         - File</span></span><br><span data-ttu-id="c6a0c-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-619">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-622">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-623">
         - Selection</span></span><br><span data-ttu-id="c6a0c-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-624">
         - Settings</span></span><br><span data-ttu-id="c6a0c-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-625">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-626">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-627">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-628">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-630">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="c6a0c-630">Office on iPad</span></span><br><span data-ttu-id="c6a0c-631">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c6a0c-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-638">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-639">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-641">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-642">
         - File</span></span><br><span data-ttu-id="c6a0c-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-644">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-647">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-648">
         - Selection</span></span><br><span data-ttu-id="c6a0c-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-649">
         - Settings</span></span><br><span data-ttu-id="c6a0c-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-650">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-651">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-652">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-653">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-655">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-655">Office on Mac</span></span><br><span data-ttu-id="c6a0c-656">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-657">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-657">- TaskPane</span></span><br><span data-ttu-id="c6a0c-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c6a0c-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-665">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-666">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-668">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-669">
         - File</span></span><br><span data-ttu-id="c6a0c-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-671">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-674">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-675">
         - Selection</span></span><br><span data-ttu-id="c6a0c-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-676">
         - Settings</span></span><br><span data-ttu-id="c6a0c-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-677">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-678">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-679">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-680">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-682">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-682">Office 2019 on Mac</span></span><br><span data-ttu-id="c6a0c-683">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-684">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="c6a0c-684">- TaskPane</span></span><br><span data-ttu-id="c6a0c-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c6a0c-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c6a0c-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c6a0c-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-691">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-692">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-694">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-695">
         - File</span></span><br><span data-ttu-id="c6a0c-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-697">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-700">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-701">
         - Selection</span></span><br><span data-ttu-id="c6a0c-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-702">
         - Settings</span></span><br><span data-ttu-id="c6a0c-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-703">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-704">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-705">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-706">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-708">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-708">Office 2016 on Mac</span></span><br><span data-ttu-id="c6a0c-709">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-710">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c6a0c-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-714">- BindingEvents</span></span><br><span data-ttu-id="c6a0c-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-715">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c6a0c-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c6a0c-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-718">
         - File</span></span><br><span data-ttu-id="c6a0c-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c6a0c-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c6a0c-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c6a0c-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-723">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-724">
         - Selection</span></span><br><span data-ttu-id="c6a0c-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-725">
         - Settings</span></span><br><span data-ttu-id="c6a0c-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-726">
         - TableBindings</span></span><br><span data-ttu-id="c6a0c-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-727">
         - TableCoercion</span></span><br><span data-ttu-id="c6a0c-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-728">
         - TextBindings</span></span><br><span data-ttu-id="c6a0c-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-729">
         - TextCoercion</span></span><br><span data-ttu-id="c6a0c-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c6a0c-731">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c6a0c-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c6a0c-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c6a0c-733">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-733">Platform</span></span></th>
    <th><span data-ttu-id="c6a0c-734">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-734">Extension points</span></span></th>
    <th><span data-ttu-id="c6a0c-735">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="c6a0c-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-737">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="c6a0c-738">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-738">- Content</span></span><br><span data-ttu-id="c6a0c-739">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-739">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-745">- ActiveView</span></span><br><span data-ttu-id="c6a0c-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-746">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-747">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-748">
         - File</span></span><br><span data-ttu-id="c6a0c-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-749">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-750">
         - Selection</span></span><br><span data-ttu-id="c6a0c-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-751">
         - Settings</span></span><br><span data-ttu-id="c6a0c-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-753">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-753">Office on Windows</span></span><br><span data-ttu-id="c6a0c-754">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-755">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-755">- Content</span></span><br><span data-ttu-id="c6a0c-756">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-756">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-762">- ActiveView</span></span><br><span data-ttu-id="c6a0c-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-763">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-764">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-765">
         - File</span></span><br><span data-ttu-id="c6a0c-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-766">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-767">
         - Selection</span></span><br><span data-ttu-id="c6a0c-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-768">
         - Settings</span></span><br><span data-ttu-id="c6a0c-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-770">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-770">Office 2019 on Windows</span></span><br><span data-ttu-id="c6a0c-771">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-772">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-772">- Content</span></span><br><span data-ttu-id="c6a0c-773">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-773">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-777">- ActiveView</span></span><br><span data-ttu-id="c6a0c-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-778">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-779">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-780">
         - File</span></span><br><span data-ttu-id="c6a0c-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-781">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-782">
         - Selection</span></span><br><span data-ttu-id="c6a0c-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-783">
         - Settings</span></span><br><span data-ttu-id="c6a0c-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-785">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-785">Office 2016 on Windows</span></span><br><span data-ttu-id="c6a0c-786">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-787">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-787">- Content</span></span><br><span data-ttu-id="c6a0c-788">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c6a0c-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-791">- ActiveView</span></span><br><span data-ttu-id="c6a0c-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-792">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-793">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-794">
         - File</span></span><br><span data-ttu-id="c6a0c-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-795">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-796">
         - Selection</span></span><br><span data-ttu-id="c6a0c-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-797">
         - Settings</span></span><br><span data-ttu-id="c6a0c-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-799">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-799">Office 2013 on Windows</span></span><br><span data-ttu-id="c6a0c-800">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-801">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-801">- Content</span></span><br><span data-ttu-id="c6a0c-802">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c6a0c-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c6a0c-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-805">- ActiveView</span></span><br><span data-ttu-id="c6a0c-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-806">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-807">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-808">
         - File</span></span><br><span data-ttu-id="c6a0c-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-809">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-810">
         - Selection</span></span><br><span data-ttu-id="c6a0c-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-811">
         - Settings</span></span><br><span data-ttu-id="c6a0c-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-813">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="c6a0c-813">Office on iPad</span></span><br><span data-ttu-id="c6a0c-814">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-815">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-815">- Content</span></span><br><span data-ttu-id="c6a0c-816">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-820">- ActiveView</span></span><br><span data-ttu-id="c6a0c-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-821">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-822">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-823">
         - File</span></span><br><span data-ttu-id="c6a0c-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-824">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-825">
         - Selection</span></span><br><span data-ttu-id="c6a0c-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-826">
         - Settings</span></span><br><span data-ttu-id="c6a0c-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-828">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-828">Office on Mac</span></span><br><span data-ttu-id="c6a0c-829">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c6a0c-830">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-830">- Content</span></span><br><span data-ttu-id="c6a0c-831">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-831">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c6a0c-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-837">- ActiveView</span></span><br><span data-ttu-id="c6a0c-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-838">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-839">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-840">
         - File</span></span><br><span data-ttu-id="c6a0c-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-841">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-842">
         - Selection</span></span><br><span data-ttu-id="c6a0c-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-843">
         - Settings</span></span><br><span data-ttu-id="c6a0c-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-845">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-845">Office 2019 on Mac</span></span><br><span data-ttu-id="c6a0c-846">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-847">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-847">- Content</span></span><br><span data-ttu-id="c6a0c-848">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-848">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-852">- ActiveView</span></span><br><span data-ttu-id="c6a0c-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-853">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-854">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-855">
         - File</span></span><br><span data-ttu-id="c6a0c-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-856">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-857">
         - Selection</span></span><br><span data-ttu-id="c6a0c-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-858">
         - Settings</span></span><br><span data-ttu-id="c6a0c-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-860">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-860">Office 2016 on Mac</span></span><br><span data-ttu-id="c6a0c-861">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-862">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-862">- Content</span></span><br><span data-ttu-id="c6a0c-863">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c6a0c-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c6a0c-866">- ActiveView</span></span><br><span data-ttu-id="c6a0c-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-867">
         - CompressedFile</span></span><br><span data-ttu-id="c6a0c-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-868">
         - DocumentEvents</span></span><br><span data-ttu-id="c6a0c-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="c6a0c-869">
         - File</span></span><br><span data-ttu-id="c6a0c-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c6a0c-870">
         - PdfFile</span></span><br><span data-ttu-id="c6a0c-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-871">
         - Selection</span></span><br><span data-ttu-id="c6a0c-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-872">
         - Settings</span></span><br><span data-ttu-id="c6a0c-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c6a0c-874">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="c6a0c-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c6a0c-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="c6a0c-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c6a0c-876">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-876">Platform</span></span></th>
    <th><span data-ttu-id="c6a0c-877">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-877">Extension points</span></span></th>
    <th><span data-ttu-id="c6a0c-878">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="c6a0c-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-880">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c6a0c-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="c6a0c-881">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c6a0c-881">- Content</span></span><br><span data-ttu-id="c6a0c-882">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-882">
         - TaskPane</span></span><br><span data-ttu-id="c6a0c-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c6a0c-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c6a0c-887">- DocumentEvents</span></span><br><span data-ttu-id="c6a0c-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="c6a0c-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c6a0c-889">
         - Settings</span></span><br><span data-ttu-id="c6a0c-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c6a0c-891">Projet</span><span class="sxs-lookup"><span data-stu-id="c6a0c-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c6a0c-892">Plateforme</span><span class="sxs-lookup"><span data-stu-id="c6a0c-892">Platform</span></span></th>
    <th><span data-ttu-id="c6a0c-893">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c6a0c-893">Extension points</span></span></th>
    <th><span data-ttu-id="c6a0c-894">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c6a0c-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="c6a0c-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-896">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-896">Office 2019 on Windows</span></span><br><span data-ttu-id="c6a0c-897">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-898">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-900">- Selection</span></span><br><span data-ttu-id="c6a0c-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-902">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-902">Office 2016 on Windows</span></span><br><span data-ttu-id="c6a0c-903">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-904">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-906">- Selection</span></span><br><span data-ttu-id="c6a0c-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c6a0c-908">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="c6a0c-908">Office 2013 on Windows</span></span><br><span data-ttu-id="c6a0c-909">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c6a0c-910">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c6a0c-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c6a0c-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c6a0c-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="c6a0c-912">- Selection</span></span><br><span data-ttu-id="c6a0c-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c6a0c-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c6a0c-914">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c6a0c-914">See also</span></span>

- [<span data-ttu-id="c6a0c-915">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c6a0c-916">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="c6a0c-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c6a0c-917">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="c6a0c-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c6a0c-918">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="c6a0c-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c6a0c-919">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="c6a0c-919">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c6a0c-920">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="c6a0c-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c6a0c-921">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c6a0c-922">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c6a0c-923">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c6a0c-924">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c6a0c-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c6a0c-925">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="c6a0c-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
