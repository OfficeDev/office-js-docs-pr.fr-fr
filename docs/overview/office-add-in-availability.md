---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: 510f2419d5d364a536f8c96f2057505161f03993
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804645"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e6a71-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="e6a71-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e6a71-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="e6a71-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e6a71-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="e6a71-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="e6a71-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="e6a71-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="e6a71-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e6a71-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e6a71-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e6a71-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e6a71-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e6a71-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e6a71-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="e6a71-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-114">- TaskPane</span></span><br><span data-ttu-id="e6a71-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-115">
        - Content</span></span><br><span data-ttu-id="e6a71-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-116">
        - Custom Functions</span></span><br><span data-ttu-id="e6a71-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="e6a71-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e6a71-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e6a71-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e6a71-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-130">
        - BindingEvents</span></span><br><span data-ttu-id="e6a71-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-131">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-132">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-133">
        - File</span></span><br><span data-ttu-id="e6a71-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-134">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-136">
        - Selection</span></span><br><span data-ttu-id="e6a71-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-137">
        - Settings</span></span><br><span data-ttu-id="e6a71-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-138">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-139">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-140">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-142">Office on Windows</span></span><br><span data-ttu-id="e6a71-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-144">- TaskPane</span></span><br><span data-ttu-id="e6a71-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-145">
        - Content</span></span><br><span data-ttu-id="e6a71-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-146">
        - Custom Functions</span></span><br><span data-ttu-id="e6a71-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="e6a71-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e6a71-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e6a71-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e6a71-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-160">
        - BindingEvents</span></span><br><span data-ttu-id="e6a71-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-161">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-162">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-163">
        - File</span></span><br><span data-ttu-id="e6a71-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-164">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-166">
        - Selection</span></span><br><span data-ttu-id="e6a71-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-167">
        - Settings</span></span><br><span data-ttu-id="e6a71-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-168">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-169">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-170">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-172">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-172">Office 2019 on Windows</span></span><br><span data-ttu-id="e6a71-173">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e6a71-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-174">- TaskPane</span></span><br><span data-ttu-id="e6a71-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-175">
        - Content</span></span><br><span data-ttu-id="e6a71-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e6a71-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-187">- BindingEvents</span></span><br><span data-ttu-id="e6a71-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-188">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-189">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-190">
        - File</span></span><br><span data-ttu-id="e6a71-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-191">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-193">
        - Selection</span></span><br><span data-ttu-id="e6a71-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-194">
        - Settings</span></span><br><span data-ttu-id="e6a71-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-195">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-196">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-197">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-199">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-199">Office 2016 on Windows</span></span><br><span data-ttu-id="e6a71-200">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e6a71-201">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-201">- TaskPane</span></span><br><span data-ttu-id="e6a71-202">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-202">
        - Content</span></span></td>
    <td><span data-ttu-id="e6a71-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e6a71-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-206">- BindingEvents</span></span><br><span data-ttu-id="e6a71-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-207">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-208">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-209">
        - File</span></span><br><span data-ttu-id="e6a71-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-210">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-212">
        - Selection</span></span><br><span data-ttu-id="e6a71-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-213">
        - Settings</span></span><br><span data-ttu-id="e6a71-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-214">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-215">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-216">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-218">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-218">Office 2013 on Windows</span></span><br><span data-ttu-id="e6a71-219">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e6a71-220">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-220">
        - TaskPane</span></span><br><span data-ttu-id="e6a71-221">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e6a71-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e6a71-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e6a71-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-224">
        - BindingEvents</span></span><br><span data-ttu-id="e6a71-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-225">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-226">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-227">
        - File</span></span><br><span data-ttu-id="e6a71-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-228">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-230">
        - Selection</span></span><br><span data-ttu-id="e6a71-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-231">
        - Settings</span></span><br><span data-ttu-id="e6a71-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-232">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-233">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-234">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-236">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="e6a71-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e6a71-237">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e6a71-238">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-238">- TaskPane</span></span><br><span data-ttu-id="e6a71-239">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-239">
        - Content</span></span><br><span data-ttu-id="e6a71-240">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e6a71-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e6a71-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-252">- BindingEvents</span></span><br><span data-ttu-id="e6a71-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-253">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-254">
        - File</span></span><br><span data-ttu-id="e6a71-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-255">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-257">
        - Selection</span></span><br><span data-ttu-id="e6a71-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-258">
        - Settings</span></span><br><span data-ttu-id="e6a71-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-259">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-260">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-261">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-263">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-263">Office apps on Mac</span></span><br><span data-ttu-id="e6a71-264">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e6a71-265">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-265">- TaskPane</span></span><br><span data-ttu-id="e6a71-266">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-266">
        - Content</span></span><br><span data-ttu-id="e6a71-267">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-267">
        - Custom Functions</span></span><br><span data-ttu-id="e6a71-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e6a71-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e6a71-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e6a71-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-281">- BindingEvents</span></span><br><span data-ttu-id="e6a71-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-282">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-283">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-284">
        - File</span></span><br><span data-ttu-id="e6a71-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-285">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-287">
        - PdfFile</span></span><br><span data-ttu-id="e6a71-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-288">
        - Selection</span></span><br><span data-ttu-id="e6a71-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-289">
        - Settings</span></span><br><span data-ttu-id="e6a71-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-290">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-291">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-292">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-294">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-294">Office 2019 for Mac</span></span><br><span data-ttu-id="e6a71-295">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e6a71-296">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-296">- TaskPane</span></span><br><span data-ttu-id="e6a71-297">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-297">
        - Content</span></span><br><span data-ttu-id="e6a71-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e6a71-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e6a71-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e6a71-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e6a71-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e6a71-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e6a71-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e6a71-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e6a71-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-309">- BindingEvents</span></span><br><span data-ttu-id="e6a71-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-310">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-311">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-312">
        - File</span></span><br><span data-ttu-id="e6a71-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-313">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-315">
        - PdfFile</span></span><br><span data-ttu-id="e6a71-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-316">
        - Selection</span></span><br><span data-ttu-id="e6a71-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-317">
        - Settings</span></span><br><span data-ttu-id="e6a71-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-318">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-319">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-320">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-322">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e6a71-323">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e6a71-324">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-324">- TaskPane</span></span><br><span data-ttu-id="e6a71-325">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-325">
        - Content</span></span></td>
    <td><span data-ttu-id="e6a71-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e6a71-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e6a71-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e6a71-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-329">- BindingEvents</span></span><br><span data-ttu-id="e6a71-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-330">
        - CompressedFile</span></span><br><span data-ttu-id="e6a71-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-331">
        - DocumentEvents</span></span><br><span data-ttu-id="e6a71-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-332">
        - File</span></span><br><span data-ttu-id="e6a71-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-333">
        - MatrixBindings</span></span><br><span data-ttu-id="e6a71-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-335">
        - PdfFile</span></span><br><span data-ttu-id="e6a71-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-336">
        - Selection</span></span><br><span data-ttu-id="e6a71-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-337">
        - Settings</span></span><br><span data-ttu-id="e6a71-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-338">
        - TableBindings</span></span><br><span data-ttu-id="e6a71-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-339">
        - TableCoercion</span></span><br><span data-ttu-id="e6a71-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-340">
        - TextBindings</span></span><br><span data-ttu-id="e6a71-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e6a71-342">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e6a71-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="e6a71-343">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e6a71-344">Plateforme</span><span class="sxs-lookup"><span data-stu-id="e6a71-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e6a71-345">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e6a71-346">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e6a71-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-348">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-348">Office on the web</span></span></td>
    <td><span data-ttu-id="e6a71-349">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e6a71-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-351">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-351">Office on Windows</span></span><br><span data-ttu-id="e6a71-352">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e6a71-353">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e6a71-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-355">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-355">Office for Mac</span></span><br><span data-ttu-id="e6a71-356">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="e6a71-357">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e6a71-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e6a71-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="e6a71-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="e6a71-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e6a71-360">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e6a71-360">Platform</span></span></th>
    <th><span data-ttu-id="e6a71-361">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-361">Extension points</span></span></th>
    <th><span data-ttu-id="e6a71-362">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="e6a71-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-364">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-364">Office on the web</span></span><br><span data-ttu-id="e6a71-365">(moderne)</span><span class="sxs-lookup"><span data-stu-id="e6a71-365">Modern</span></span></td>
    <td> <span data-ttu-id="e6a71-366">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-366">- Mail Read</span></span><br><span data-ttu-id="e6a71-367">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-367">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e6a71-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e6a71-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-377">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-377">Office on the web</span></span><br><span data-ttu-id="e6a71-378">(classique)</span><span class="sxs-lookup"><span data-stu-id="e6a71-378">Classic.</span></span></td>
    <td> <span data-ttu-id="e6a71-379">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-379">- Mail Read</span></span><br><span data-ttu-id="e6a71-380">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-380">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e6a71-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-389">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-389">Office on Windows</span></span><br><span data-ttu-id="e6a71-390">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-391">- Mail Read</span></span><br><span data-ttu-id="e6a71-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-392">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e6a71-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="e6a71-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e6a71-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e6a71-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e6a71-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-403">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-403">Office 2019 on Windows</span></span><br><span data-ttu-id="e6a71-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-405">- Mail Read</span></span><br><span data-ttu-id="e6a71-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-406">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e6a71-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="e6a71-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e6a71-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e6a71-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e6a71-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-417">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-417">Office 2016 on Windows</span></span><br><span data-ttu-id="e6a71-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-419">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-419">- Mail Read</span></span><br><span data-ttu-id="e6a71-420">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-420">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e6a71-422">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="e6a71-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e6a71-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="e6a71-427">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-428">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-428">Office 2013 on Windows</span></span><br><span data-ttu-id="e6a71-429">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-430">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-430">- Mail Read</span></span><br><span data-ttu-id="e6a71-431">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="e6a71-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="e6a71-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="e6a71-436">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-437">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="e6a71-437">Office apps on iOS</span></span><br><span data-ttu-id="e6a71-438">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-439">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-439">- Mail Read</span></span><br><span data-ttu-id="e6a71-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e6a71-446">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-447">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-447">Office apps on Mac</span></span><br><span data-ttu-id="e6a71-448">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-449">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-449">- Mail Read</span></span><br><span data-ttu-id="e6a71-450">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-450">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e6a71-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e6a71-459">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-460">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-460">Office 2019 for Mac</span></span><br><span data-ttu-id="e6a71-461">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-462">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-462">- Mail Read</span></span><br><span data-ttu-id="e6a71-463">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-463">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e6a71-471">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-472">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e6a71-473">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-474">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-474">- Mail Read</span></span><br><span data-ttu-id="e6a71-475">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-475">
      - Mail Compose</span></span><br><span data-ttu-id="e6a71-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e6a71-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e6a71-483">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-484">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="e6a71-484">Office apps on Android</span></span><br><span data-ttu-id="e6a71-485">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-486">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e6a71-486">- Mail Read</span></span><br><span data-ttu-id="e6a71-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e6a71-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e6a71-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e6a71-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e6a71-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e6a71-493">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e6a71-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="e6a71-494">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e6a71-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="e6a71-495">Word</span><span class="sxs-lookup"><span data-stu-id="e6a71-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e6a71-496">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e6a71-496">Platform</span></span></th>
    <th><span data-ttu-id="e6a71-497">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-497">Extension points</span></span></th>
    <th><span data-ttu-id="e6a71-498">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="e6a71-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-500">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="e6a71-501">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-501">- TaskPane</span></span><br><span data-ttu-id="e6a71-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e6a71-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-509">- BindingEvents</span></span><br><span data-ttu-id="e6a71-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-511">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-512">
         - File</span></span><br><span data-ttu-id="e6a71-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-514">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-517">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-518">
         - Selection</span></span><br><span data-ttu-id="e6a71-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-519">
         - Settings</span></span><br><span data-ttu-id="e6a71-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-520">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-521">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-522">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-523">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-525">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-525">Office on Windows</span></span><br><span data-ttu-id="e6a71-526">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-527">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-527">- TaskPane</span></span><br><span data-ttu-id="e6a71-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e6a71-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-535">- BindingEvents</span></span><br><span data-ttu-id="e6a71-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-536">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-538">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-539">
         - File</span></span><br><span data-ttu-id="e6a71-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-541">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-544">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-545">
         - Selection</span></span><br><span data-ttu-id="e6a71-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-546">
         - Settings</span></span><br><span data-ttu-id="e6a71-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-547">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-548">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-549">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-550">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-552">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-552">Office 2019 on Windows</span></span><br><span data-ttu-id="e6a71-553">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-554">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="e6a71-554">- TaskPane</span></span><br><span data-ttu-id="e6a71-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-561">- BindingEvents</span></span><br><span data-ttu-id="e6a71-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-562">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-564">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-565">
         - File</span></span><br><span data-ttu-id="e6a71-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-567">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-570">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-571">
         - Selection</span></span><br><span data-ttu-id="e6a71-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-572">
         - Settings</span></span><br><span data-ttu-id="e6a71-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-573">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-574">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-575">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-576">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-578">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-578">Office 2016 on Windows</span></span><br><span data-ttu-id="e6a71-579">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-580">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e6a71-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-584">- BindingEvents</span></span><br><span data-ttu-id="e6a71-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-585">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-587">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-588">
         - File</span></span><br><span data-ttu-id="e6a71-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-590">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-593">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-594">
         - Selection</span></span><br><span data-ttu-id="e6a71-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-595">
         - Settings</span></span><br><span data-ttu-id="e6a71-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-596">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-597">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-598">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-599">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-601">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-601">Office 2013 on Windows</span></span><br><span data-ttu-id="e6a71-602">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-603">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e6a71-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e6a71-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-606">- BindingEvents</span></span><br><span data-ttu-id="e6a71-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-607">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-609">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-610">
         - File</span></span><br><span data-ttu-id="e6a71-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-612">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-615">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-616">
         - Selection</span></span><br><span data-ttu-id="e6a71-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-617">
         - Settings</span></span><br><span data-ttu-id="e6a71-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-618">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-619">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-620">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-621">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-623">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="e6a71-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e6a71-624">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-625">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="e6a71-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-631">- BindingEvents</span></span><br><span data-ttu-id="e6a71-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-632">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-634">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-635">
         - File</span></span><br><span data-ttu-id="e6a71-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-637">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-640">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-641">
         - Selection</span></span><br><span data-ttu-id="e6a71-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-642">
         - Settings</span></span><br><span data-ttu-id="e6a71-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-643">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-644">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-645">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-646">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-648">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-648">Office apps on Mac</span></span><br><span data-ttu-id="e6a71-649">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-650">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-650">- TaskPane</span></span><br><span data-ttu-id="e6a71-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="e6a71-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-658">- BindingEvents</span></span><br><span data-ttu-id="e6a71-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-659">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-661">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-662">
         - File</span></span><br><span data-ttu-id="e6a71-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-664">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-667">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-668">
         - Selection</span></span><br><span data-ttu-id="e6a71-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-669">
         - Settings</span></span><br><span data-ttu-id="e6a71-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-670">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-671">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-672">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-673">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-675">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-675">Office 2019 for Mac</span></span><br><span data-ttu-id="e6a71-676">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-677">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="e6a71-677">- TaskPane</span></span><br><span data-ttu-id="e6a71-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="e6a71-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="e6a71-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="e6a71-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-684">- BindingEvents</span></span><br><span data-ttu-id="e6a71-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-685">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-687">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-688">
         - File</span></span><br><span data-ttu-id="e6a71-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-690">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-693">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-694">
         - Selection</span></span><br><span data-ttu-id="e6a71-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-695">
         - Settings</span></span><br><span data-ttu-id="e6a71-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-696">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-697">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-698">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-699">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-701">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e6a71-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="e6a71-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e6a71-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e6a71-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-707">- BindingEvents</span></span><br><span data-ttu-id="e6a71-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-708">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e6a71-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="e6a71-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-710">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-711">
         - File</span></span><br><span data-ttu-id="e6a71-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-713">
         - MatrixBindings</span></span><br><span data-ttu-id="e6a71-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="e6a71-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e6a71-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-716">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-717">
         - Selection</span></span><br><span data-ttu-id="e6a71-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-718">
         - Settings</span></span><br><span data-ttu-id="e6a71-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-719">
         - TableBindings</span></span><br><span data-ttu-id="e6a71-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-720">
         - TableCoercion</span></span><br><span data-ttu-id="e6a71-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e6a71-721">
         - TextBindings</span></span><br><span data-ttu-id="e6a71-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-722">
         - TextCoercion</span></span><br><span data-ttu-id="e6a71-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="e6a71-724">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e6a71-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e6a71-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e6a71-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e6a71-726">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e6a71-726">Platform</span></span></th>
    <th><span data-ttu-id="e6a71-727">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-727">Extension points</span></span></th>
    <th><span data-ttu-id="e6a71-728">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="e6a71-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-730">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="e6a71-731">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-731">- Content</span></span><br><span data-ttu-id="e6a71-732">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-732">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e6a71-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-737">- ActiveView</span></span><br><span data-ttu-id="e6a71-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-738">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-739">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-740">
         - File</span></span><br><span data-ttu-id="e6a71-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-741">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-742">
         - Selection</span></span><br><span data-ttu-id="e6a71-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-743">
         - Settings</span></span><br><span data-ttu-id="e6a71-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-745">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-745">Office on Windows</span></span><br><span data-ttu-id="e6a71-746">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-747">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-747">- Content</span></span><br><span data-ttu-id="e6a71-748">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-748">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e6a71-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-753">- ActiveView</span></span><br><span data-ttu-id="e6a71-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-754">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-755">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-756">
         - File</span></span><br><span data-ttu-id="e6a71-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-757">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-758">
         - Selection</span></span><br><span data-ttu-id="e6a71-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-759">
         - Settings</span></span><br><span data-ttu-id="e6a71-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-761">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-761">Office 2019 on Windows</span></span><br><span data-ttu-id="e6a71-762">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-763">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-763">- Content</span></span><br><span data-ttu-id="e6a71-764">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-764">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-768">- ActiveView</span></span><br><span data-ttu-id="e6a71-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-769">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-770">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-771">
         - File</span></span><br><span data-ttu-id="e6a71-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-772">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-773">
         - Selection</span></span><br><span data-ttu-id="e6a71-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-774">
         - Settings</span></span><br><span data-ttu-id="e6a71-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-776">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-776">Office 2016 on Windows</span></span><br><span data-ttu-id="e6a71-777">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-778">- Content</span></span><br><span data-ttu-id="e6a71-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e6a71-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e6a71-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-782">- ActiveView</span></span><br><span data-ttu-id="e6a71-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-783">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-784">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-785">
         - File</span></span><br><span data-ttu-id="e6a71-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-786">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-787">
         - Selection</span></span><br><span data-ttu-id="e6a71-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-788">
         - Settings</span></span><br><span data-ttu-id="e6a71-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-790">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-790">Office 2013 on Windows</span></span><br><span data-ttu-id="e6a71-791">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-792">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-792">- Content</span></span><br><span data-ttu-id="e6a71-793">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e6a71-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e6a71-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e6a71-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-796">- ActiveView</span></span><br><span data-ttu-id="e6a71-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-797">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-798">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-799">
         - File</span></span><br><span data-ttu-id="e6a71-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-800">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-801">
         - Selection</span></span><br><span data-ttu-id="e6a71-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-802">
         - Settings</span></span><br><span data-ttu-id="e6a71-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-804">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="e6a71-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e6a71-805">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-806">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-806">- Content</span></span><br><span data-ttu-id="e6a71-807">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-810">- ActiveView</span></span><br><span data-ttu-id="e6a71-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-811">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-812">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-813">
         - File</span></span><br><span data-ttu-id="e6a71-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-814">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-815">
         - Selection</span></span><br><span data-ttu-id="e6a71-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-816">
         - Settings</span></span><br><span data-ttu-id="e6a71-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-818">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-818">Office apps on Mac</span></span><br><span data-ttu-id="e6a71-819">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="e6a71-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e6a71-820">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-820">- Content</span></span><br><span data-ttu-id="e6a71-821">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-821">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e6a71-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e6a71-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-826">- ActiveView</span></span><br><span data-ttu-id="e6a71-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-827">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-828">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-829">
         - File</span></span><br><span data-ttu-id="e6a71-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-830">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-831">
         - Selection</span></span><br><span data-ttu-id="e6a71-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-832">
         - Settings</span></span><br><span data-ttu-id="e6a71-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-834">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-834">Office 2019 for Mac</span></span><br><span data-ttu-id="e6a71-835">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-836">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-836">- Content</span></span><br><span data-ttu-id="e6a71-837">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-837">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-841">- ActiveView</span></span><br><span data-ttu-id="e6a71-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-842">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-843">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-844">
         - File</span></span><br><span data-ttu-id="e6a71-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-845">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-846">
         - Selection</span></span><br><span data-ttu-id="e6a71-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-847">
         - Settings</span></span><br><span data-ttu-id="e6a71-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-849">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e6a71-850">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-851">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-851">- Content</span></span><br><span data-ttu-id="e6a71-852">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e6a71-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e6a71-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e6a71-855">- ActiveView</span></span><br><span data-ttu-id="e6a71-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-856">
         - CompressedFile</span></span><br><span data-ttu-id="e6a71-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-857">
         - DocumentEvents</span></span><br><span data-ttu-id="e6a71-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="e6a71-858">
         - File</span></span><br><span data-ttu-id="e6a71-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e6a71-859">
         - PdfFile</span></span><br><span data-ttu-id="e6a71-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-860">
         - Selection</span></span><br><span data-ttu-id="e6a71-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-861">
         - Settings</span></span><br><span data-ttu-id="e6a71-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e6a71-863">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e6a71-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="e6a71-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="e6a71-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e6a71-865">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e6a71-865">Platform</span></span></th>
    <th><span data-ttu-id="e6a71-866">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-866">Extension points</span></span></th>
    <th><span data-ttu-id="e6a71-867">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="e6a71-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-869">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e6a71-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="e6a71-870">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e6a71-870">- Content</span></span><br><span data-ttu-id="e6a71-871">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-871">
         - TaskPane</span></span><br><span data-ttu-id="e6a71-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e6a71-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e6a71-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e6a71-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e6a71-876">- DocumentEvents</span></span><br><span data-ttu-id="e6a71-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="e6a71-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e6a71-878">
         - Settings</span></span><br><span data-ttu-id="e6a71-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e6a71-880">Projet</span><span class="sxs-lookup"><span data-stu-id="e6a71-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e6a71-881">Plateforme</span><span class="sxs-lookup"><span data-stu-id="e6a71-881">Platform</span></span></th>
    <th><span data-ttu-id="e6a71-882">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e6a71-882">Extension points</span></span></th>
    <th><span data-ttu-id="e6a71-883">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e6a71-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="e6a71-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e6a71-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-885">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-885">Office 2019 on Windows</span></span><br><span data-ttu-id="e6a71-886">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-887">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-889">- Selection</span></span><br><span data-ttu-id="e6a71-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-891">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-891">Office 2016 on Windows</span></span><br><span data-ttu-id="e6a71-892">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-893">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-895">- Selection</span></span><br><span data-ttu-id="e6a71-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e6a71-897">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="e6a71-897">Office 2013 on Windows</span></span><br><span data-ttu-id="e6a71-898">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e6a71-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e6a71-899">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e6a71-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e6a71-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e6a71-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="e6a71-901">- Selection</span></span><br><span data-ttu-id="e6a71-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e6a71-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e6a71-903">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e6a71-903">See also</span></span>

- [<span data-ttu-id="e6a71-904">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e6a71-905">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e6a71-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="e6a71-906">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="e6a71-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e6a71-907">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="e6a71-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e6a71-908">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="e6a71-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="e6a71-909">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="e6a71-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="e6a71-910">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="e6a71-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="e6a71-911">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="e6a71-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="e6a71-912">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="e6a71-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="e6a71-913">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="e6a71-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="e6a71-914">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="e6a71-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
