---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/26/2019
localization_priority: Priority
ms.openlocfilehash: 7039ca59af22f1101bdff7b6bcd4506497d6c9cd
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940835"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f2cca-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="f2cca-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f2cca-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="f2cca-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="f2cca-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="f2cca-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="f2cca-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="f2cca-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="f2cca-108">Excel</span><span class="sxs-lookup"><span data-stu-id="f2cca-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f2cca-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f2cca-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f2cca-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f2cca-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f2cca-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="f2cca-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-114">- TaskPane</span></span><br><span data-ttu-id="f2cca-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-115">
        - Content</span></span><br><span data-ttu-id="f2cca-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-116">
        - Custom Functions</span></span><br><span data-ttu-id="f2cca-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="f2cca-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f2cca-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f2cca-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f2cca-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-130">
        - BindingEvents</span></span><br><span data-ttu-id="f2cca-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-131">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-132">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-133">
        - File</span></span><br><span data-ttu-id="f2cca-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-134">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-136">
        - Selection</span></span><br><span data-ttu-id="f2cca-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-137">
        - Settings</span></span><br><span data-ttu-id="f2cca-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-138">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-139">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-140">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-142">Office on Windows</span></span><br><span data-ttu-id="f2cca-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-144">- TaskPane</span></span><br><span data-ttu-id="f2cca-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-145">
        - Content</span></span><br><span data-ttu-id="f2cca-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-146">
        - Custom Functions</span></span><br><span data-ttu-id="f2cca-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="f2cca-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f2cca-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f2cca-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f2cca-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-160">
        - BindingEvents</span></span><br><span data-ttu-id="f2cca-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-161">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-162">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-163">
        - File</span></span><br><span data-ttu-id="f2cca-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-164">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-166">
        - Selection</span></span><br><span data-ttu-id="f2cca-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-167">
        - Settings</span></span><br><span data-ttu-id="f2cca-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-168">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-169">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-170">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-172">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-172">Office 2019 on Windows</span></span><br><span data-ttu-id="f2cca-173">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f2cca-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-174">- TaskPane</span></span><br><span data-ttu-id="f2cca-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-175">
        - Content</span></span><br><span data-ttu-id="f2cca-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f2cca-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-187">- BindingEvents</span></span><br><span data-ttu-id="f2cca-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-188">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-189">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-190">
        - File</span></span><br><span data-ttu-id="f2cca-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-191">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-193">
        - Selection</span></span><br><span data-ttu-id="f2cca-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-194">
        - Settings</span></span><br><span data-ttu-id="f2cca-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-195">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-196">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-197">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-199">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-199">Office 2016 on Windows</span></span><br><span data-ttu-id="f2cca-200">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f2cca-201">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-201">- TaskPane</span></span><br><span data-ttu-id="f2cca-202">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-202">
        - Content</span></span></td>
    <td><span data-ttu-id="f2cca-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f2cca-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-206">- BindingEvents</span></span><br><span data-ttu-id="f2cca-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-207">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-208">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-209">
        - File</span></span><br><span data-ttu-id="f2cca-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-210">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-212">
        - Selection</span></span><br><span data-ttu-id="f2cca-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-213">
        - Settings</span></span><br><span data-ttu-id="f2cca-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-214">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-215">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-216">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-218">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-218">Office 2013 on Windows</span></span><br><span data-ttu-id="f2cca-219">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f2cca-220">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-220">
        - TaskPane</span></span><br><span data-ttu-id="f2cca-221">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="f2cca-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f2cca-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f2cca-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-224">
        - BindingEvents</span></span><br><span data-ttu-id="f2cca-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-225">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-226">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-227">
        - File</span></span><br><span data-ttu-id="f2cca-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-228">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-230">
        - Selection</span></span><br><span data-ttu-id="f2cca-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-231">
        - Settings</span></span><br><span data-ttu-id="f2cca-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-232">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-233">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-234">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-236">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="f2cca-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f2cca-237">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f2cca-238">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-238">- TaskPane</span></span><br><span data-ttu-id="f2cca-239">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-239">
        - Content</span></span><br><span data-ttu-id="f2cca-240">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f2cca-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f2cca-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-252">- BindingEvents</span></span><br><span data-ttu-id="f2cca-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-253">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-254">
        - File</span></span><br><span data-ttu-id="f2cca-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-255">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-257">
        - Selection</span></span><br><span data-ttu-id="f2cca-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-258">
        - Settings</span></span><br><span data-ttu-id="f2cca-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-259">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-260">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-261">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-263">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-263">Office apps on Mac</span></span><br><span data-ttu-id="f2cca-264">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f2cca-265">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-265">- TaskPane</span></span><br><span data-ttu-id="f2cca-266">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-266">
        - Content</span></span><br><span data-ttu-id="f2cca-267">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-267">
        - Custom Functions</span></span><br><span data-ttu-id="f2cca-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f2cca-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f2cca-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f2cca-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-281">- BindingEvents</span></span><br><span data-ttu-id="f2cca-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-282">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-283">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-284">
        - File</span></span><br><span data-ttu-id="f2cca-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-285">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-287">
        - PdfFile</span></span><br><span data-ttu-id="f2cca-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-288">
        - Selection</span></span><br><span data-ttu-id="f2cca-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-289">
        - Settings</span></span><br><span data-ttu-id="f2cca-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-290">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-291">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-292">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-294">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-294">Office 2019 for Mac</span></span><br><span data-ttu-id="f2cca-295">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f2cca-296">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-296">- TaskPane</span></span><br><span data-ttu-id="f2cca-297">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-297">
        - Content</span></span><br><span data-ttu-id="f2cca-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f2cca-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f2cca-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f2cca-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f2cca-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f2cca-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f2cca-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f2cca-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f2cca-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-309">- BindingEvents</span></span><br><span data-ttu-id="f2cca-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-310">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-311">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-312">
        - File</span></span><br><span data-ttu-id="f2cca-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-313">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-315">
        - PdfFile</span></span><br><span data-ttu-id="f2cca-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-316">
        - Selection</span></span><br><span data-ttu-id="f2cca-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-317">
        - Settings</span></span><br><span data-ttu-id="f2cca-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-318">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-319">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-320">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-322">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="f2cca-323">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f2cca-324">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-324">- TaskPane</span></span><br><span data-ttu-id="f2cca-325">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-325">
        - Content</span></span></td>
    <td><span data-ttu-id="f2cca-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f2cca-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f2cca-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f2cca-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-329">- BindingEvents</span></span><br><span data-ttu-id="f2cca-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-330">
        - CompressedFile</span></span><br><span data-ttu-id="f2cca-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-331">
        - DocumentEvents</span></span><br><span data-ttu-id="f2cca-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-332">
        - File</span></span><br><span data-ttu-id="f2cca-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-333">
        - MatrixBindings</span></span><br><span data-ttu-id="f2cca-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-335">
        - PdfFile</span></span><br><span data-ttu-id="f2cca-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-336">
        - Selection</span></span><br><span data-ttu-id="f2cca-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-337">
        - Settings</span></span><br><span data-ttu-id="f2cca-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-338">
        - TableBindings</span></span><br><span data-ttu-id="f2cca-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-339">
        - TableCoercion</span></span><br><span data-ttu-id="f2cca-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-340">
        - TextBindings</span></span><br><span data-ttu-id="f2cca-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f2cca-342">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="f2cca-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="f2cca-343">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f2cca-344">Plateforme</span><span class="sxs-lookup"><span data-stu-id="f2cca-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f2cca-345">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f2cca-346">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f2cca-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-348">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-348">Office on the web</span></span></td>
    <td><span data-ttu-id="f2cca-349">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f2cca-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-351">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-351">Office on Windows</span></span><br><span data-ttu-id="f2cca-352">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f2cca-353">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f2cca-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-355">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-355">Office for Mac</span></span><br><span data-ttu-id="f2cca-356">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="f2cca-357">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f2cca-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f2cca-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="f2cca-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="f2cca-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cca-360">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f2cca-360">Platform</span></span></th>
    <th><span data-ttu-id="f2cca-361">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-361">Extension points</span></span></th>
    <th><span data-ttu-id="f2cca-362">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="f2cca-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-364">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-364">Office on the web</span></span><br><span data-ttu-id="f2cca-365">(moderne)</span><span class="sxs-lookup"><span data-stu-id="f2cca-365">Modern</span></span></td>
    <td> <span data-ttu-id="f2cca-366">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-366">- Mail Read</span></span><br><span data-ttu-id="f2cca-367">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-367">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f2cca-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f2cca-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-377">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-377">Office on the web</span></span><br><span data-ttu-id="f2cca-378">(classique)</span><span class="sxs-lookup"><span data-stu-id="f2cca-378">Classic.</span></span></td>
    <td> <span data-ttu-id="f2cca-379">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-379">- Mail Read</span></span><br><span data-ttu-id="f2cca-380">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-380">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f2cca-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-389">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-389">Office on Windows</span></span><br><span data-ttu-id="f2cca-390">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-391">- Mail Read</span></span><br><span data-ttu-id="f2cca-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-392">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f2cca-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="f2cca-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f2cca-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f2cca-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f2cca-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-403">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-403">Office 2019 on Windows</span></span><br><span data-ttu-id="f2cca-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-405">- Mail Read</span></span><br><span data-ttu-id="f2cca-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-406">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f2cca-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="f2cca-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f2cca-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f2cca-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f2cca-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-417">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-417">Office 2016 on Windows</span></span><br><span data-ttu-id="f2cca-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-419">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-419">- Mail Read</span></span><br><span data-ttu-id="f2cca-420">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-420">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f2cca-422">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="f2cca-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f2cca-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f2cca-427">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-428">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-428">Office 2013 on Windows</span></span><br><span data-ttu-id="f2cca-429">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-430">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-430">- Mail Read</span></span><br><span data-ttu-id="f2cca-431">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="f2cca-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="f2cca-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f2cca-436">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-437">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="f2cca-437">Office apps on iOS</span></span><br><span data-ttu-id="f2cca-438">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-439">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-439">- Mail Read</span></span><br><span data-ttu-id="f2cca-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f2cca-446">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-447">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-447">Office apps on Mac</span></span><br><span data-ttu-id="f2cca-448">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-449">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-449">- Mail Read</span></span><br><span data-ttu-id="f2cca-450">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-450">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f2cca-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f2cca-459">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-460">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-460">Office 2019 for Mac</span></span><br><span data-ttu-id="f2cca-461">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-462">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-462">- Mail Read</span></span><br><span data-ttu-id="f2cca-463">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-463">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f2cca-471">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-472">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="f2cca-473">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-474">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-474">- Mail Read</span></span><br><span data-ttu-id="f2cca-475">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-475">
      - Mail Compose</span></span><br><span data-ttu-id="f2cca-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f2cca-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f2cca-483">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-484">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="f2cca-484">Office apps on Android</span></span><br><span data-ttu-id="f2cca-485">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-486">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f2cca-486">- Mail Read</span></span><br><span data-ttu-id="f2cca-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f2cca-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f2cca-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f2cca-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f2cca-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f2cca-493">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f2cca-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="f2cca-494">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="f2cca-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="f2cca-495">Word</span><span class="sxs-lookup"><span data-stu-id="f2cca-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cca-496">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f2cca-496">Platform</span></span></th>
    <th><span data-ttu-id="f2cca-497">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-497">Extension points</span></span></th>
    <th><span data-ttu-id="f2cca-498">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="f2cca-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-500">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="f2cca-501">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-501">- TaskPane</span></span><br><span data-ttu-id="f2cca-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f2cca-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-509">- BindingEvents</span></span><br><span data-ttu-id="f2cca-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-511">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-512">
         - File</span></span><br><span data-ttu-id="f2cca-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-514">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-517">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-518">
         - Selection</span></span><br><span data-ttu-id="f2cca-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-519">
         - Settings</span></span><br><span data-ttu-id="f2cca-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-520">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-521">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-522">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-523">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-525">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-525">Office on Windows</span></span><br><span data-ttu-id="f2cca-526">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-527">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-527">- TaskPane</span></span><br><span data-ttu-id="f2cca-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f2cca-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-535">- BindingEvents</span></span><br><span data-ttu-id="f2cca-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-536">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-538">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-539">
         - File</span></span><br><span data-ttu-id="f2cca-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-541">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-544">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-545">
         - Selection</span></span><br><span data-ttu-id="f2cca-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-546">
         - Settings</span></span><br><span data-ttu-id="f2cca-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-547">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-548">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-549">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-550">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-552">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-552">Office 2019 on Windows</span></span><br><span data-ttu-id="f2cca-553">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-554">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="f2cca-554">- TaskPane</span></span><br><span data-ttu-id="f2cca-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-561">- BindingEvents</span></span><br><span data-ttu-id="f2cca-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-562">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-564">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-565">
         - File</span></span><br><span data-ttu-id="f2cca-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-567">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-570">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-571">
         - Selection</span></span><br><span data-ttu-id="f2cca-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-572">
         - Settings</span></span><br><span data-ttu-id="f2cca-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-573">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-574">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-575">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-576">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-578">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-578">Office 2016 on Windows</span></span><br><span data-ttu-id="f2cca-579">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-580">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f2cca-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-584">- BindingEvents</span></span><br><span data-ttu-id="f2cca-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-585">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-587">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-588">
         - File</span></span><br><span data-ttu-id="f2cca-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-590">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-593">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-594">
         - Selection</span></span><br><span data-ttu-id="f2cca-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-595">
         - Settings</span></span><br><span data-ttu-id="f2cca-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-596">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-597">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-598">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-599">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-601">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-601">Office 2013 on Windows</span></span><br><span data-ttu-id="f2cca-602">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-603">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f2cca-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f2cca-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-606">- BindingEvents</span></span><br><span data-ttu-id="f2cca-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-607">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-609">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-610">
         - File</span></span><br><span data-ttu-id="f2cca-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-612">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-615">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-616">
         - Selection</span></span><br><span data-ttu-id="f2cca-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-617">
         - Settings</span></span><br><span data-ttu-id="f2cca-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-618">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-619">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-620">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-621">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-623">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="f2cca-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f2cca-624">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-625">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f2cca-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-631">- BindingEvents</span></span><br><span data-ttu-id="f2cca-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-632">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-634">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-635">
         - File</span></span><br><span data-ttu-id="f2cca-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-637">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-640">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-641">
         - Selection</span></span><br><span data-ttu-id="f2cca-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-642">
         - Settings</span></span><br><span data-ttu-id="f2cca-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-643">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-644">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-645">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-646">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-648">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-648">Office apps on Mac</span></span><br><span data-ttu-id="f2cca-649">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-650">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-650">- TaskPane</span></span><br><span data-ttu-id="f2cca-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="f2cca-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-658">- BindingEvents</span></span><br><span data-ttu-id="f2cca-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-659">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-661">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-662">
         - File</span></span><br><span data-ttu-id="f2cca-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-664">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-667">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-668">
         - Selection</span></span><br><span data-ttu-id="f2cca-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-669">
         - Settings</span></span><br><span data-ttu-id="f2cca-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-670">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-671">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-672">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-673">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-675">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-675">Office 2019 for Mac</span></span><br><span data-ttu-id="f2cca-676">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-677">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="f2cca-677">- TaskPane</span></span><br><span data-ttu-id="f2cca-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f2cca-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f2cca-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f2cca-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-684">- BindingEvents</span></span><br><span data-ttu-id="f2cca-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-685">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-687">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-688">
         - File</span></span><br><span data-ttu-id="f2cca-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-690">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-693">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-694">
         - Selection</span></span><br><span data-ttu-id="f2cca-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-695">
         - Settings</span></span><br><span data-ttu-id="f2cca-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-696">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-697">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-698">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-699">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-701">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="f2cca-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f2cca-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f2cca-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f2cca-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-707">- BindingEvents</span></span><br><span data-ttu-id="f2cca-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-708">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cca-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="f2cca-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-710">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-711">
         - File</span></span><br><span data-ttu-id="f2cca-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-713">
         - MatrixBindings</span></span><br><span data-ttu-id="f2cca-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="f2cca-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f2cca-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-716">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-717">
         - Selection</span></span><br><span data-ttu-id="f2cca-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-718">
         - Settings</span></span><br><span data-ttu-id="f2cca-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-719">
         - TableBindings</span></span><br><span data-ttu-id="f2cca-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-720">
         - TableCoercion</span></span><br><span data-ttu-id="f2cca-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cca-721">
         - TextBindings</span></span><br><span data-ttu-id="f2cca-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-722">
         - TextCoercion</span></span><br><span data-ttu-id="f2cca-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="f2cca-724">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="f2cca-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f2cca-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f2cca-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cca-726">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f2cca-726">Platform</span></span></th>
    <th><span data-ttu-id="f2cca-727">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-727">Extension points</span></span></th>
    <th><span data-ttu-id="f2cca-728">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="f2cca-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-730">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="f2cca-731">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-731">- Content</span></span><br><span data-ttu-id="f2cca-732">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-732">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f2cca-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f2cca-738">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-738">- ActiveView</span></span><br><span data-ttu-id="f2cca-739">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-739">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-740">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-741">
         - File</span></span><br><span data-ttu-id="f2cca-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-742">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-743">
         - Selection</span></span><br><span data-ttu-id="f2cca-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-744">
         - Settings</span></span><br><span data-ttu-id="f2cca-745">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-745">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-746">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-746">Office on Windows</span></span><br><span data-ttu-id="f2cca-747">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-747">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-748">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-748">- Content</span></span><br><span data-ttu-id="f2cca-749">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-749">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f2cca-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f2cca-755">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-755">- ActiveView</span></span><br><span data-ttu-id="f2cca-756">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-756">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-757">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-757">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-758">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-758">
         - File</span></span><br><span data-ttu-id="f2cca-759">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-759">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-760">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-760">
         - Selection</span></span><br><span data-ttu-id="f2cca-761">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-761">
         - Settings</span></span><br><span data-ttu-id="f2cca-762">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-762">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-763">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-763">Office 2019 on Windows</span></span><br><span data-ttu-id="f2cca-764">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-764">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-765">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-765">- Content</span></span><br><span data-ttu-id="f2cca-766">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-766">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-770">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-770">- ActiveView</span></span><br><span data-ttu-id="f2cca-771">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-771">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-772">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-772">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-773">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-773">
         - File</span></span><br><span data-ttu-id="f2cca-774">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-774">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-775">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-775">
         - Selection</span></span><br><span data-ttu-id="f2cca-776">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-776">
         - Settings</span></span><br><span data-ttu-id="f2cca-777">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-777">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-778">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-778">Office 2016 on Windows</span></span><br><span data-ttu-id="f2cca-779">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-779">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-780">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-780">- Content</span></span><br><span data-ttu-id="f2cca-781">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-781">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f2cca-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f2cca-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-784">- ActiveView</span></span><br><span data-ttu-id="f2cca-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-785">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-786">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-787">
         - File</span></span><br><span data-ttu-id="f2cca-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-788">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-789">
         - Selection</span></span><br><span data-ttu-id="f2cca-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-790">
         - Settings</span></span><br><span data-ttu-id="f2cca-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-792">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-792">Office 2013 on Windows</span></span><br><span data-ttu-id="f2cca-793">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-794">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-794">- Content</span></span><br><span data-ttu-id="f2cca-795">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-795">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="f2cca-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f2cca-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f2cca-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-798">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-798">- ActiveView</span></span><br><span data-ttu-id="f2cca-799">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-799">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-800">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-800">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-801">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-801">
         - File</span></span><br><span data-ttu-id="f2cca-802">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-802">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-803">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-803">
         - Selection</span></span><br><span data-ttu-id="f2cca-804">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-804">
         - Settings</span></span><br><span data-ttu-id="f2cca-805">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-805">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-806">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="f2cca-806">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f2cca-807">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-807">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-808">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-808">- Content</span></span><br><span data-ttu-id="f2cca-809">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-809">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f2cca-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-813">- ActiveView</span></span><br><span data-ttu-id="f2cca-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-814">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-815">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-816">
         - File</span></span><br><span data-ttu-id="f2cca-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-817">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-818">
         - Selection</span></span><br><span data-ttu-id="f2cca-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-819">
         - Settings</span></span><br><span data-ttu-id="f2cca-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-821">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-821">Office apps on Mac</span></span><br><span data-ttu-id="f2cca-822">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f2cca-822">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f2cca-823">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-823">- Content</span></span><br><span data-ttu-id="f2cca-824">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-824">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f2cca-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f2cca-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f2cca-830">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-830">- ActiveView</span></span><br><span data-ttu-id="f2cca-831">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-831">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-832">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-832">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-833">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-833">
         - File</span></span><br><span data-ttu-id="f2cca-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-834">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-835">
         - Selection</span></span><br><span data-ttu-id="f2cca-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-836">
         - Settings</span></span><br><span data-ttu-id="f2cca-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-838">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-838">Office 2019 for Mac</span></span><br><span data-ttu-id="f2cca-839">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-840">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-840">- Content</span></span><br><span data-ttu-id="f2cca-841">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-841">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-845">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-845">- ActiveView</span></span><br><span data-ttu-id="f2cca-846">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-846">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-847">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-847">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-848">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-848">
         - File</span></span><br><span data-ttu-id="f2cca-849">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-849">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-850">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-850">
         - Selection</span></span><br><span data-ttu-id="f2cca-851">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-851">
         - Settings</span></span><br><span data-ttu-id="f2cca-852">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-852">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-853">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-853">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="f2cca-854">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-854">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-855">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-855">- Content</span></span><br><span data-ttu-id="f2cca-856">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-856">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f2cca-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f2cca-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cca-859">- ActiveView</span></span><br><span data-ttu-id="f2cca-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-860">
         - CompressedFile</span></span><br><span data-ttu-id="f2cca-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-861">
         - DocumentEvents</span></span><br><span data-ttu-id="f2cca-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="f2cca-862">
         - File</span></span><br><span data-ttu-id="f2cca-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cca-863">
         - PdfFile</span></span><br><span data-ttu-id="f2cca-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-864">
         - Selection</span></span><br><span data-ttu-id="f2cca-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-865">
         - Settings</span></span><br><span data-ttu-id="f2cca-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-866">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f2cca-867">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="f2cca-867">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="f2cca-868">OneNote</span><span class="sxs-lookup"><span data-stu-id="f2cca-868">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cca-869">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f2cca-869">Platform</span></span></th>
    <th><span data-ttu-id="f2cca-870">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-870">Extension points</span></span></th>
    <th><span data-ttu-id="f2cca-871">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-871">API requirement sets</span></span></th>
    <th><span data-ttu-id="f2cca-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-873">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f2cca-873">Office on the web</span></span></td>
    <td> <span data-ttu-id="f2cca-874">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f2cca-874">- Content</span></span><br><span data-ttu-id="f2cca-875">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-875">
         - TaskPane</span></span><br><span data-ttu-id="f2cca-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f2cca-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="f2cca-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f2cca-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-880">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cca-880">- DocumentEvents</span></span><br><span data-ttu-id="f2cca-881">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-881">
         - HtmlCoercion</span></span><br><span data-ttu-id="f2cca-882">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f2cca-882">
         - Settings</span></span><br><span data-ttu-id="f2cca-883">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-883">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="f2cca-884">Projet</span><span class="sxs-lookup"><span data-stu-id="f2cca-884">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cca-885">Plateforme</span><span class="sxs-lookup"><span data-stu-id="f2cca-885">Platform</span></span></th>
    <th><span data-ttu-id="f2cca-886">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f2cca-886">Extension points</span></span></th>
    <th><span data-ttu-id="f2cca-887">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f2cca-887">API requirement sets</span></span></th>
    <th><span data-ttu-id="f2cca-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f2cca-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-889">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-889">Office 2019 on Windows</span></span><br><span data-ttu-id="f2cca-890">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-890">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-891">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-891">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-893">- Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-893">- Selection</span></span><br><span data-ttu-id="f2cca-894">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-894">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-895">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-895">Office 2016 on Windows</span></span><br><span data-ttu-id="f2cca-896">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-896">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-897">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-897">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-899">- Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-899">- Selection</span></span><br><span data-ttu-id="f2cca-900">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-900">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cca-901">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="f2cca-901">Office 2013 on Windows</span></span><br><span data-ttu-id="f2cca-902">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="f2cca-902">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f2cca-903">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-903">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f2cca-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f2cca-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f2cca-905">- Selection</span><span class="sxs-lookup"><span data-stu-id="f2cca-905">- Selection</span></span><br><span data-ttu-id="f2cca-906">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cca-906">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f2cca-907">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f2cca-907">See also</span></span>

- [<span data-ttu-id="f2cca-908">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-908">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f2cca-909">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2cca-909">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="f2cca-910">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="f2cca-910">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="f2cca-911">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="f2cca-911">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="f2cca-912">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="f2cca-912">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="f2cca-913">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="f2cca-913">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="f2cca-914">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="f2cca-914">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="f2cca-915">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="f2cca-915">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="f2cca-916">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="f2cca-916">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="f2cca-917">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="f2cca-917">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="f2cca-918">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="f2cca-918">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
