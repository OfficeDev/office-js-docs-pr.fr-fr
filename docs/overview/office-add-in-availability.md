---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: 956ee6b8a9e990a3d6d942ee4a65a1e9275ea025
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851368"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="580b6-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="580b6-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="580b6-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="580b6-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="580b6-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="580b6-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="580b6-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="580b6-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="580b6-108">Excel</span><span class="sxs-lookup"><span data-stu-id="580b6-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="580b6-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="580b6-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="580b6-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="580b6-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="580b6-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="580b6-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-114">- TaskPane</span></span><br><span data-ttu-id="580b6-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-115">
        - Content</span></span><br><span data-ttu-id="580b6-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-116">
        - Custom Functions</span></span><br><span data-ttu-id="580b6-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="580b6-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="580b6-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="580b6-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="580b6-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="580b6-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="580b6-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="580b6-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="580b6-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-130">
        - BindingEvents</span></span><br><span data-ttu-id="580b6-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-131">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-132">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-133">
        - File</span></span><br><span data-ttu-id="580b6-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-134">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-136">
        - Selection</span></span><br><span data-ttu-id="580b6-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-137">
        - Settings</span></span><br><span data-ttu-id="580b6-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-138">
        - TableBindings</span></span><br><span data-ttu-id="580b6-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-139">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-140">
        - TextBindings</span></span><br><span data-ttu-id="580b6-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-142">Office on Windows</span></span><br><span data-ttu-id="580b6-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-144">- TaskPane</span></span><br><span data-ttu-id="580b6-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-145">
        - Content</span></span><br><span data-ttu-id="580b6-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-146">
        - Custom Functions</span></span><br><span data-ttu-id="580b6-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="580b6-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="580b6-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="580b6-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="580b6-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="580b6-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="580b6-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="580b6-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-161">
        - BindingEvents</span></span><br><span data-ttu-id="580b6-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-162">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-163">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-164">
        - File</span></span><br><span data-ttu-id="580b6-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-165">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-167">
        - Selection</span></span><br><span data-ttu-id="580b6-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-168">
        - Settings</span></span><br><span data-ttu-id="580b6-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-169">
        - TableBindings</span></span><br><span data-ttu-id="580b6-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-170">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-171">
        - TextBindings</span></span><br><span data-ttu-id="580b6-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-173">Office 2019 on Windows</span></span><br><span data-ttu-id="580b6-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="580b6-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-175">- TaskPane</span></span><br><span data-ttu-id="580b6-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-176">
        - Content</span></span><br><span data-ttu-id="580b6-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="580b6-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-188">- BindingEvents</span></span><br><span data-ttu-id="580b6-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-189">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-190">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-191">
        - File</span></span><br><span data-ttu-id="580b6-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-192">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-194">
        - Selection</span></span><br><span data-ttu-id="580b6-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-195">
        - Settings</span></span><br><span data-ttu-id="580b6-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-196">
        - TableBindings</span></span><br><span data-ttu-id="580b6-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-197">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-198">
        - TextBindings</span></span><br><span data-ttu-id="580b6-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-200">Office 2016 on Windows</span></span><br><span data-ttu-id="580b6-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="580b6-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-202">- TaskPane</span></span><br><span data-ttu-id="580b6-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-203">
        - Content</span></span></td>
    <td><span data-ttu-id="580b6-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="580b6-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-207">- BindingEvents</span></span><br><span data-ttu-id="580b6-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-208">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-209">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-210">
        - File</span></span><br><span data-ttu-id="580b6-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-211">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-213">
        - Selection</span></span><br><span data-ttu-id="580b6-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-214">
        - Settings</span></span><br><span data-ttu-id="580b6-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-215">
        - TableBindings</span></span><br><span data-ttu-id="580b6-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-216">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-217">
        - TextBindings</span></span><br><span data-ttu-id="580b6-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-219">Office 2013 on Windows</span></span><br><span data-ttu-id="580b6-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="580b6-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-221">
        - TaskPane</span></span><br><span data-ttu-id="580b6-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="580b6-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="580b6-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="580b6-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-225">
        - BindingEvents</span></span><br><span data-ttu-id="580b6-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-226">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-227">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-228">
        - File</span></span><br><span data-ttu-id="580b6-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-229">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-231">
        - Selection</span></span><br><span data-ttu-id="580b6-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-232">
        - Settings</span></span><br><span data-ttu-id="580b6-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-233">
        - TableBindings</span></span><br><span data-ttu-id="580b6-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-234">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-235">
        - TextBindings</span></span><br><span data-ttu-id="580b6-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="580b6-237">Office on iPad</span></span><br><span data-ttu-id="580b6-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="580b6-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-239">- TaskPane</span></span><br><span data-ttu-id="580b6-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-240">
        - Content</span></span></td>
    <td><span data-ttu-id="580b6-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="580b6-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="580b6-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="580b6-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="580b6-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-253">- BindingEvents</span></span><br><span data-ttu-id="580b6-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-254">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-255">
        - File</span></span><br><span data-ttu-id="580b6-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-256">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-258">
        - Selection</span></span><br><span data-ttu-id="580b6-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-259">
        - Settings</span></span><br><span data-ttu-id="580b6-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-260">
        - TableBindings</span></span><br><span data-ttu-id="580b6-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-261">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-262">
        - TextBindings</span></span><br><span data-ttu-id="580b6-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-264">Office on Mac</span></span><br><span data-ttu-id="580b6-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="580b6-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-266">- TaskPane</span></span><br><span data-ttu-id="580b6-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-267">
        - Content</span></span><br><span data-ttu-id="580b6-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-268">
        - Custom Functions</span></span><br><span data-ttu-id="580b6-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="580b6-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="580b6-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="580b6-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="580b6-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="580b6-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="580b6-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-283">- BindingEvents</span></span><br><span data-ttu-id="580b6-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-284">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-285">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-286">
        - File</span></span><br><span data-ttu-id="580b6-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-287">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-289">
        - PdfFile</span></span><br><span data-ttu-id="580b6-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-290">
        - Selection</span></span><br><span data-ttu-id="580b6-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-291">
        - Settings</span></span><br><span data-ttu-id="580b6-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-292">
        - TableBindings</span></span><br><span data-ttu-id="580b6-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-293">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-294">
        - TextBindings</span></span><br><span data-ttu-id="580b6-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-296">Office 2019 on Mac</span></span><br><span data-ttu-id="580b6-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="580b6-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-298">- TaskPane</span></span><br><span data-ttu-id="580b6-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-299">
        - Content</span></span><br><span data-ttu-id="580b6-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="580b6-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="580b6-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="580b6-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="580b6-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="580b6-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="580b6-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="580b6-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="580b6-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-311">- BindingEvents</span></span><br><span data-ttu-id="580b6-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-312">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-313">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-314">
        - File</span></span><br><span data-ttu-id="580b6-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-315">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-317">
        - PdfFile</span></span><br><span data-ttu-id="580b6-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-318">
        - Selection</span></span><br><span data-ttu-id="580b6-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-319">
        - Settings</span></span><br><span data-ttu-id="580b6-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-320">
        - TableBindings</span></span><br><span data-ttu-id="580b6-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-321">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-322">
        - TextBindings</span></span><br><span data-ttu-id="580b6-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-324">Office 2016 on Mac</span></span><br><span data-ttu-id="580b6-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="580b6-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-326">- TaskPane</span></span><br><span data-ttu-id="580b6-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-327">
        - Content</span></span></td>
    <td><span data-ttu-id="580b6-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="580b6-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="580b6-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="580b6-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-331">- BindingEvents</span></span><br><span data-ttu-id="580b6-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-332">
        - CompressedFile</span></span><br><span data-ttu-id="580b6-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-333">
        - DocumentEvents</span></span><br><span data-ttu-id="580b6-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="580b6-334">
        - File</span></span><br><span data-ttu-id="580b6-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-335">
        - MatrixBindings</span></span><br><span data-ttu-id="580b6-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="580b6-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-337">
        - PdfFile</span></span><br><span data-ttu-id="580b6-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-338">
        - Selection</span></span><br><span data-ttu-id="580b6-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-339">
        - Settings</span></span><br><span data-ttu-id="580b6-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-340">
        - TableBindings</span></span><br><span data-ttu-id="580b6-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-341">
        - TableCoercion</span></span><br><span data-ttu-id="580b6-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-342">
        - TextBindings</span></span><br><span data-ttu-id="580b6-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="580b6-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="580b6-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="580b6-345">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="580b6-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="580b6-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="580b6-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="580b6-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="580b6-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-350">Office on the web</span></span></td>
    <td><span data-ttu-id="580b6-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="580b6-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-353">Office on Windows</span></span><br><span data-ttu-id="580b6-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="580b6-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="580b6-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-357">Office for Mac</span></span><br><span data-ttu-id="580b6-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="580b6-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="580b6-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="580b6-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="580b6-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="580b6-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="580b6-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="580b6-362">Platform</span></span></th>
    <th><span data-ttu-id="580b6-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-363">Extension points</span></span></th>
    <th><span data-ttu-id="580b6-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="580b6-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-366">Office on the web</span></span><br><span data-ttu-id="580b6-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="580b6-367">(modern)</span></span></td>
    <td> <span data-ttu-id="580b6-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-368">- Mail Read</span></span><br><span data-ttu-id="580b6-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-369">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="580b6-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="580b6-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="580b6-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-380">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-380">Office on the web</span></span><br><span data-ttu-id="580b6-381">(classique)</span><span class="sxs-lookup"><span data-stu-id="580b6-381">(classic)</span></span></td>
    <td> <span data-ttu-id="580b6-382">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-382">- Mail Read</span></span><br><span data-ttu-id="580b6-383">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-383">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="580b6-391">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-392">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-392">Office on Windows</span></span><br><span data-ttu-id="580b6-393">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-394">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-394">- Mail Read</span></span><br><span data-ttu-id="580b6-395">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-395">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="580b6-397">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="580b6-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="580b6-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="580b6-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="580b6-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="580b6-406">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-407">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-407">Office 2019 on Windows</span></span><br><span data-ttu-id="580b6-408">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-409">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-409">- Mail Read</span></span><br><span data-ttu-id="580b6-410">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-410">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="580b6-412">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="580b6-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="580b6-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="580b6-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="580b6-420">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-421">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-421">Office 2016 on Windows</span></span><br><span data-ttu-id="580b6-422">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-423">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-423">- Mail Read</span></span><br><span data-ttu-id="580b6-424">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-424">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="580b6-426">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="580b6-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="580b6-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="580b6-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-432">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-432">Office 2013 on Windows</span></span><br><span data-ttu-id="580b6-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-434">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-434">- Mail Read</span></span><br><span data-ttu-id="580b6-435">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="580b6-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="580b6-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="580b6-440">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-441">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="580b6-441">Office on iOS</span></span><br><span data-ttu-id="580b6-442">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-443">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-443">- Mail Read</span></span><br><span data-ttu-id="580b6-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="580b6-450">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-451">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-451">Office on Mac</span></span><br><span data-ttu-id="580b6-452">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-453">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-453">- Mail Read</span></span><br><span data-ttu-id="580b6-454">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-454">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="580b6-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="580b6-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="580b6-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="580b6-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="580b6-464">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-465">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-465">Office 2019 on Mac</span></span><br><span data-ttu-id="580b6-466">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-467">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-467">- Mail Read</span></span><br><span data-ttu-id="580b6-468">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-468">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="580b6-476">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-477">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-477">Office 2016 on Mac</span></span><br><span data-ttu-id="580b6-478">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-479">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-479">- Mail Read</span></span><br><span data-ttu-id="580b6-480">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="580b6-480">
      - Mail Compose</span></span><br><span data-ttu-id="580b6-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="580b6-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="580b6-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="580b6-488">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-489">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="580b6-489">Office on Android</span></span><br><span data-ttu-id="580b6-490">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-491">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="580b6-491">- Mail Read</span></span><br><span data-ttu-id="580b6-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="580b6-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="580b6-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="580b6-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="580b6-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="580b6-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="580b6-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="580b6-498">Non disponible</span><span class="sxs-lookup"><span data-stu-id="580b6-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="580b6-499">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="580b6-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="580b6-500">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="580b6-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="580b6-501">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="580b6-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="580b6-502">Word</span><span class="sxs-lookup"><span data-stu-id="580b6-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="580b6-503">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="580b6-503">Platform</span></span></th>
    <th><span data-ttu-id="580b6-504">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-504">Extension points</span></span></th>
    <th><span data-ttu-id="580b6-505">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="580b6-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-507">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="580b6-508">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-508">- TaskPane</span></span><br><span data-ttu-id="580b6-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="580b6-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-516">- BindingEvents</span></span><br><span data-ttu-id="580b6-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-518">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-519">
         - File</span></span><br><span data-ttu-id="580b6-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-521">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-524">
         - PdfFile</span></span><br><span data-ttu-id="580b6-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-525">
         - Selection</span></span><br><span data-ttu-id="580b6-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-526">
         - Settings</span></span><br><span data-ttu-id="580b6-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-527">
         - TableBindings</span></span><br><span data-ttu-id="580b6-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-528">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-529">
         - TextBindings</span></span><br><span data-ttu-id="580b6-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-530">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-532">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-532">Office on Windows</span></span><br><span data-ttu-id="580b6-533">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-534">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-534">- TaskPane</span></span><br><span data-ttu-id="580b6-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="580b6-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-542">- BindingEvents</span></span><br><span data-ttu-id="580b6-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-543">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-545">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-546">
         - File</span></span><br><span data-ttu-id="580b6-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-548">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-551">
         - PdfFile</span></span><br><span data-ttu-id="580b6-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-552">
         - Selection</span></span><br><span data-ttu-id="580b6-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-553">
         - Settings</span></span><br><span data-ttu-id="580b6-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-554">
         - TableBindings</span></span><br><span data-ttu-id="580b6-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-555">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-556">
         - TextBindings</span></span><br><span data-ttu-id="580b6-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-557">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-559">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-559">Office 2019 on Windows</span></span><br><span data-ttu-id="580b6-560">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-561">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="580b6-561">- TaskPane</span></span><br><span data-ttu-id="580b6-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-568">- BindingEvents</span></span><br><span data-ttu-id="580b6-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-569">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-571">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-572">
         - File</span></span><br><span data-ttu-id="580b6-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-574">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-577">
         - PdfFile</span></span><br><span data-ttu-id="580b6-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-578">
         - Selection</span></span><br><span data-ttu-id="580b6-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-579">
         - Settings</span></span><br><span data-ttu-id="580b6-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-580">
         - TableBindings</span></span><br><span data-ttu-id="580b6-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-581">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-582">
         - TextBindings</span></span><br><span data-ttu-id="580b6-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-583">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-585">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-585">Office 2016 on Windows</span></span><br><span data-ttu-id="580b6-586">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-587">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="580b6-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-591">- BindingEvents</span></span><br><span data-ttu-id="580b6-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-592">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-594">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-595">
         - File</span></span><br><span data-ttu-id="580b6-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-597">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-600">
         - PdfFile</span></span><br><span data-ttu-id="580b6-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-601">
         - Selection</span></span><br><span data-ttu-id="580b6-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-602">
         - Settings</span></span><br><span data-ttu-id="580b6-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-603">
         - TableBindings</span></span><br><span data-ttu-id="580b6-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-604">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-605">
         - TextBindings</span></span><br><span data-ttu-id="580b6-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-606">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-608">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-608">Office 2013 on Windows</span></span><br><span data-ttu-id="580b6-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="580b6-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="580b6-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-613">- BindingEvents</span></span><br><span data-ttu-id="580b6-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-614">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-616">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-617">
         - File</span></span><br><span data-ttu-id="580b6-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-619">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-622">
         - PdfFile</span></span><br><span data-ttu-id="580b6-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-623">
         - Selection</span></span><br><span data-ttu-id="580b6-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-624">
         - Settings</span></span><br><span data-ttu-id="580b6-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-625">
         - TableBindings</span></span><br><span data-ttu-id="580b6-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-626">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-627">
         - TextBindings</span></span><br><span data-ttu-id="580b6-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-628">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-630">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="580b6-630">Office on iPad</span></span><br><span data-ttu-id="580b6-631">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="580b6-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-638">- BindingEvents</span></span><br><span data-ttu-id="580b6-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-639">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-641">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-642">
         - File</span></span><br><span data-ttu-id="580b6-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-644">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-647">
         - PdfFile</span></span><br><span data-ttu-id="580b6-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-648">
         - Selection</span></span><br><span data-ttu-id="580b6-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-649">
         - Settings</span></span><br><span data-ttu-id="580b6-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-650">
         - TableBindings</span></span><br><span data-ttu-id="580b6-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-651">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-652">
         - TextBindings</span></span><br><span data-ttu-id="580b6-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-653">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-655">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-655">Office on Mac</span></span><br><span data-ttu-id="580b6-656">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-657">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-657">- TaskPane</span></span><br><span data-ttu-id="580b6-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="580b6-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-665">- BindingEvents</span></span><br><span data-ttu-id="580b6-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-666">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-668">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-669">
         - File</span></span><br><span data-ttu-id="580b6-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-671">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-674">
         - PdfFile</span></span><br><span data-ttu-id="580b6-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-675">
         - Selection</span></span><br><span data-ttu-id="580b6-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-676">
         - Settings</span></span><br><span data-ttu-id="580b6-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-677">
         - TableBindings</span></span><br><span data-ttu-id="580b6-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-678">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-679">
         - TextBindings</span></span><br><span data-ttu-id="580b6-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-680">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-682">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-682">Office 2019 on Mac</span></span><br><span data-ttu-id="580b6-683">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-684">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="580b6-684">- TaskPane</span></span><br><span data-ttu-id="580b6-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="580b6-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="580b6-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="580b6-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="580b6-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-691">- BindingEvents</span></span><br><span data-ttu-id="580b6-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-692">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-694">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-695">
         - File</span></span><br><span data-ttu-id="580b6-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-697">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-700">
         - PdfFile</span></span><br><span data-ttu-id="580b6-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-701">
         - Selection</span></span><br><span data-ttu-id="580b6-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-702">
         - Settings</span></span><br><span data-ttu-id="580b6-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-703">
         - TableBindings</span></span><br><span data-ttu-id="580b6-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-704">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-705">
         - TextBindings</span></span><br><span data-ttu-id="580b6-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-706">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-708">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-708">Office 2016 on Mac</span></span><br><span data-ttu-id="580b6-709">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-710">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="580b6-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="580b6-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="580b6-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-714">- BindingEvents</span></span><br><span data-ttu-id="580b6-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-715">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="580b6-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="580b6-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-717">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-718">
         - File</span></span><br><span data-ttu-id="580b6-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-720">
         - MatrixBindings</span></span><br><span data-ttu-id="580b6-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="580b6-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="580b6-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-723">
         - PdfFile</span></span><br><span data-ttu-id="580b6-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-724">
         - Selection</span></span><br><span data-ttu-id="580b6-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-725">
         - Settings</span></span><br><span data-ttu-id="580b6-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-726">
         - TableBindings</span></span><br><span data-ttu-id="580b6-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-727">
         - TableCoercion</span></span><br><span data-ttu-id="580b6-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="580b6-728">
         - TextBindings</span></span><br><span data-ttu-id="580b6-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-729">
         - TextCoercion</span></span><br><span data-ttu-id="580b6-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="580b6-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="580b6-731">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="580b6-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="580b6-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="580b6-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="580b6-733">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="580b6-733">Platform</span></span></th>
    <th><span data-ttu-id="580b6-734">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-734">Extension points</span></span></th>
    <th><span data-ttu-id="580b6-735">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="580b6-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-737">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="580b6-738">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-738">- Content</span></span><br><span data-ttu-id="580b6-739">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-739">
         - TaskPane</span></span><br><span data-ttu-id="580b6-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="580b6-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="580b6-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-745">- ActiveView</span></span><br><span data-ttu-id="580b6-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-746">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-747">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-748">
         - File</span></span><br><span data-ttu-id="580b6-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-749">
         - PdfFile</span></span><br><span data-ttu-id="580b6-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-750">
         - Selection</span></span><br><span data-ttu-id="580b6-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-751">
         - Settings</span></span><br><span data-ttu-id="580b6-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-753">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-753">Office on Windows</span></span><br><span data-ttu-id="580b6-754">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-755">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-755">- Content</span></span><br><span data-ttu-id="580b6-756">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-756">
         - TaskPane</span></span><br><span data-ttu-id="580b6-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="580b6-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="580b6-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-762">- ActiveView</span></span><br><span data-ttu-id="580b6-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-763">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-764">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-765">
         - File</span></span><br><span data-ttu-id="580b6-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-766">
         - PdfFile</span></span><br><span data-ttu-id="580b6-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-767">
         - Selection</span></span><br><span data-ttu-id="580b6-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-768">
         - Settings</span></span><br><span data-ttu-id="580b6-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-770">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-770">Office 2019 on Windows</span></span><br><span data-ttu-id="580b6-771">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-772">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-772">- Content</span></span><br><span data-ttu-id="580b6-773">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-773">
         - TaskPane</span></span><br><span data-ttu-id="580b6-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-777">- ActiveView</span></span><br><span data-ttu-id="580b6-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-778">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-779">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-780">
         - File</span></span><br><span data-ttu-id="580b6-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-781">
         - PdfFile</span></span><br><span data-ttu-id="580b6-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-782">
         - Selection</span></span><br><span data-ttu-id="580b6-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-783">
         - Settings</span></span><br><span data-ttu-id="580b6-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-785">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-785">Office 2016 on Windows</span></span><br><span data-ttu-id="580b6-786">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-787">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-787">- Content</span></span><br><span data-ttu-id="580b6-788">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="580b6-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="580b6-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-791">- ActiveView</span></span><br><span data-ttu-id="580b6-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-792">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-793">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-794">
         - File</span></span><br><span data-ttu-id="580b6-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-795">
         - PdfFile</span></span><br><span data-ttu-id="580b6-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-796">
         - Selection</span></span><br><span data-ttu-id="580b6-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-797">
         - Settings</span></span><br><span data-ttu-id="580b6-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-799">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-799">Office 2013 on Windows</span></span><br><span data-ttu-id="580b6-800">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-801">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-801">- Content</span></span><br><span data-ttu-id="580b6-802">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="580b6-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="580b6-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="580b6-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-805">- ActiveView</span></span><br><span data-ttu-id="580b6-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-806">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-807">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-808">
         - File</span></span><br><span data-ttu-id="580b6-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-809">
         - PdfFile</span></span><br><span data-ttu-id="580b6-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-810">
         - Selection</span></span><br><span data-ttu-id="580b6-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-811">
         - Settings</span></span><br><span data-ttu-id="580b6-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-813">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="580b6-813">Office on iPad</span></span><br><span data-ttu-id="580b6-814">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-815">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-815">- Content</span></span><br><span data-ttu-id="580b6-816">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="580b6-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-820">- ActiveView</span></span><br><span data-ttu-id="580b6-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-821">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-822">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-823">
         - File</span></span><br><span data-ttu-id="580b6-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-824">
         - PdfFile</span></span><br><span data-ttu-id="580b6-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-825">
         - Selection</span></span><br><span data-ttu-id="580b6-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-826">
         - Settings</span></span><br><span data-ttu-id="580b6-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-828">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-828">Office on Mac</span></span><br><span data-ttu-id="580b6-829">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="580b6-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="580b6-830">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-830">- Content</span></span><br><span data-ttu-id="580b6-831">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-831">
         - TaskPane</span></span><br><span data-ttu-id="580b6-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="580b6-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="580b6-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="580b6-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="580b6-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-837">- ActiveView</span></span><br><span data-ttu-id="580b6-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-838">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-839">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-840">
         - File</span></span><br><span data-ttu-id="580b6-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-841">
         - PdfFile</span></span><br><span data-ttu-id="580b6-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-842">
         - Selection</span></span><br><span data-ttu-id="580b6-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-843">
         - Settings</span></span><br><span data-ttu-id="580b6-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-845">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-845">Office 2019 on Mac</span></span><br><span data-ttu-id="580b6-846">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-847">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-847">- Content</span></span><br><span data-ttu-id="580b6-848">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-848">
         - TaskPane</span></span><br><span data-ttu-id="580b6-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-852">- ActiveView</span></span><br><span data-ttu-id="580b6-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-853">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-854">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-855">
         - File</span></span><br><span data-ttu-id="580b6-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-856">
         - PdfFile</span></span><br><span data-ttu-id="580b6-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-857">
         - Selection</span></span><br><span data-ttu-id="580b6-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-858">
         - Settings</span></span><br><span data-ttu-id="580b6-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-860">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-860">Office 2016 on Mac</span></span><br><span data-ttu-id="580b6-861">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-862">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-862">- Content</span></span><br><span data-ttu-id="580b6-863">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="580b6-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="580b6-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="580b6-866">- ActiveView</span></span><br><span data-ttu-id="580b6-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="580b6-867">
         - CompressedFile</span></span><br><span data-ttu-id="580b6-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-868">
         - DocumentEvents</span></span><br><span data-ttu-id="580b6-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="580b6-869">
         - File</span></span><br><span data-ttu-id="580b6-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="580b6-870">
         - PdfFile</span></span><br><span data-ttu-id="580b6-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-871">
         - Selection</span></span><br><span data-ttu-id="580b6-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-872">
         - Settings</span></span><br><span data-ttu-id="580b6-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="580b6-874">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="580b6-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="580b6-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="580b6-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="580b6-876">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="580b6-876">Platform</span></span></th>
    <th><span data-ttu-id="580b6-877">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-877">Extension points</span></span></th>
    <th><span data-ttu-id="580b6-878">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="580b6-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-880">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="580b6-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="580b6-881">- Contenu</span><span class="sxs-lookup"><span data-stu-id="580b6-881">- Content</span></span><br><span data-ttu-id="580b6-882">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-882">
         - TaskPane</span></span><br><span data-ttu-id="580b6-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="580b6-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="580b6-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="580b6-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="580b6-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="580b6-887">- DocumentEvents</span></span><br><span data-ttu-id="580b6-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="580b6-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="580b6-889">
         - Settings</span></span><br><span data-ttu-id="580b6-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="580b6-891">Projet</span><span class="sxs-lookup"><span data-stu-id="580b6-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="580b6-892">Plateforme</span><span class="sxs-lookup"><span data-stu-id="580b6-892">Platform</span></span></th>
    <th><span data-ttu-id="580b6-893">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="580b6-893">Extension points</span></span></th>
    <th><span data-ttu-id="580b6-894">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="580b6-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="580b6-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-896">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-896">Office 2019 on Windows</span></span><br><span data-ttu-id="580b6-897">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-898">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-900">- Selection</span></span><br><span data-ttu-id="580b6-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-902">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-902">Office 2016 on Windows</span></span><br><span data-ttu-id="580b6-903">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-904">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-906">- Selection</span></span><br><span data-ttu-id="580b6-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="580b6-908">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="580b6-908">Office 2013 on Windows</span></span><br><span data-ttu-id="580b6-909">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="580b6-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="580b6-910">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="580b6-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="580b6-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="580b6-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="580b6-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="580b6-912">- Selection</span></span><br><span data-ttu-id="580b6-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="580b6-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="580b6-914">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="580b6-914">See also</span></span>

- [<span data-ttu-id="580b6-915">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="580b6-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="580b6-916">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="580b6-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="580b6-917">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="580b6-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="580b6-918">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="580b6-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="580b6-919">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="580b6-919">API reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="580b6-920">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="580b6-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="580b6-921">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="580b6-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="580b6-922">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="580b6-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="580b6-923">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="580b6-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="580b6-924">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="580b6-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="580b6-925">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="580b6-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="580b6-926">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="580b6-926">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)