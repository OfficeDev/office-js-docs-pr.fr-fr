---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688957"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5c919-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="5c919-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5c919-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="5c919-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="5c919-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="5c919-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="5c919-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="5c919-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="5c919-108">Excel</span><span class="sxs-lookup"><span data-stu-id="5c919-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5c919-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="5c919-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5c919-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5c919-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5c919-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="5c919-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-114">- TaskPane</span></span><br><span data-ttu-id="5c919-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-115">
        - Content</span></span><br><span data-ttu-id="5c919-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-116">
        - Custom Functions</span></span><br><span data-ttu-id="5c919-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="5c919-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5c919-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5c919-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5c919-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="5c919-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="5c919-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="5c919-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="5c919-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-130">
        - BindingEvents</span></span><br><span data-ttu-id="5c919-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-131">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-132">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-133">
        - File</span></span><br><span data-ttu-id="5c919-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-134">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-136">
        - Selection</span></span><br><span data-ttu-id="5c919-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-137">
        - Settings</span></span><br><span data-ttu-id="5c919-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-138">
        - TableBindings</span></span><br><span data-ttu-id="5c919-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-139">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-140">
        - TextBindings</span></span><br><span data-ttu-id="5c919-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-142">Office on Windows</span></span><br><span data-ttu-id="5c919-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-144">- TaskPane</span></span><br><span data-ttu-id="5c919-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-145">
        - Content</span></span><br><span data-ttu-id="5c919-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-146">
        - Custom Functions</span></span><br><span data-ttu-id="5c919-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="5c919-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5c919-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5c919-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5c919-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="5c919-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="5c919-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="5c919-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-161">
        - BindingEvents</span></span><br><span data-ttu-id="5c919-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-162">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-163">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-164">
        - File</span></span><br><span data-ttu-id="5c919-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-165">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-167">
        - Selection</span></span><br><span data-ttu-id="5c919-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-168">
        - Settings</span></span><br><span data-ttu-id="5c919-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-169">
        - TableBindings</span></span><br><span data-ttu-id="5c919-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-170">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-171">
        - TextBindings</span></span><br><span data-ttu-id="5c919-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-173">Office 2019 on Windows</span></span><br><span data-ttu-id="5c919-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5c919-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-175">- TaskPane</span></span><br><span data-ttu-id="5c919-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-176">
        - Content</span></span><br><span data-ttu-id="5c919-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5c919-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-188">- BindingEvents</span></span><br><span data-ttu-id="5c919-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-189">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-190">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-191">
        - File</span></span><br><span data-ttu-id="5c919-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-192">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-194">
        - Selection</span></span><br><span data-ttu-id="5c919-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-195">
        - Settings</span></span><br><span data-ttu-id="5c919-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-196">
        - TableBindings</span></span><br><span data-ttu-id="5c919-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-197">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-198">
        - TextBindings</span></span><br><span data-ttu-id="5c919-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-200">Office 2016 on Windows</span></span><br><span data-ttu-id="5c919-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5c919-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-202">- TaskPane</span></span><br><span data-ttu-id="5c919-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-203">
        - Content</span></span></td>
    <td><span data-ttu-id="5c919-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5c919-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-207">- BindingEvents</span></span><br><span data-ttu-id="5c919-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-208">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-209">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-210">
        - File</span></span><br><span data-ttu-id="5c919-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-211">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-213">
        - Selection</span></span><br><span data-ttu-id="5c919-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-214">
        - Settings</span></span><br><span data-ttu-id="5c919-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-215">
        - TableBindings</span></span><br><span data-ttu-id="5c919-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-216">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-217">
        - TextBindings</span></span><br><span data-ttu-id="5c919-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-219">Office 2013 on Windows</span></span><br><span data-ttu-id="5c919-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5c919-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-221">
        - TaskPane</span></span><br><span data-ttu-id="5c919-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5c919-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5c919-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5c919-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-225">
        - BindingEvents</span></span><br><span data-ttu-id="5c919-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-226">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-227">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-228">
        - File</span></span><br><span data-ttu-id="5c919-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-229">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-231">
        - Selection</span></span><br><span data-ttu-id="5c919-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-232">
        - Settings</span></span><br><span data-ttu-id="5c919-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-233">
        - TableBindings</span></span><br><span data-ttu-id="5c919-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-234">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-235">
        - TextBindings</span></span><br><span data-ttu-id="5c919-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="5c919-237">Office on iPad</span></span><br><span data-ttu-id="5c919-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5c919-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-239">- TaskPane</span></span><br><span data-ttu-id="5c919-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-240">
        - Content</span></span></td>
    <td><span data-ttu-id="5c919-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5c919-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5c919-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="5c919-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="5c919-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-253">- BindingEvents</span></span><br><span data-ttu-id="5c919-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-254">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-255">
        - File</span></span><br><span data-ttu-id="5c919-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-256">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-258">
        - Selection</span></span><br><span data-ttu-id="5c919-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-259">
        - Settings</span></span><br><span data-ttu-id="5c919-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-260">
        - TableBindings</span></span><br><span data-ttu-id="5c919-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-261">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-262">
        - TextBindings</span></span><br><span data-ttu-id="5c919-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-264">Office on Mac</span></span><br><span data-ttu-id="5c919-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5c919-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-266">- TaskPane</span></span><br><span data-ttu-id="5c919-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-267">
        - Content</span></span><br><span data-ttu-id="5c919-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-268">
        - Custom Functions</span></span><br><span data-ttu-id="5c919-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5c919-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5c919-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5c919-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="5c919-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="5c919-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="5c919-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-283">- BindingEvents</span></span><br><span data-ttu-id="5c919-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-284">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-285">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-286">
        - File</span></span><br><span data-ttu-id="5c919-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-287">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-289">
        - PdfFile</span></span><br><span data-ttu-id="5c919-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-290">
        - Selection</span></span><br><span data-ttu-id="5c919-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-291">
        - Settings</span></span><br><span data-ttu-id="5c919-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-292">
        - TableBindings</span></span><br><span data-ttu-id="5c919-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-293">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-294">
        - TextBindings</span></span><br><span data-ttu-id="5c919-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-296">Office 2019 on Mac</span></span><br><span data-ttu-id="5c919-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5c919-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-298">- TaskPane</span></span><br><span data-ttu-id="5c919-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-299">
        - Content</span></span><br><span data-ttu-id="5c919-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5c919-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5c919-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5c919-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5c919-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5c919-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5c919-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5c919-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5c919-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-311">- BindingEvents</span></span><br><span data-ttu-id="5c919-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-312">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-313">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-314">
        - File</span></span><br><span data-ttu-id="5c919-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-315">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-317">
        - PdfFile</span></span><br><span data-ttu-id="5c919-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-318">
        - Selection</span></span><br><span data-ttu-id="5c919-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-319">
        - Settings</span></span><br><span data-ttu-id="5c919-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-320">
        - TableBindings</span></span><br><span data-ttu-id="5c919-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-321">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-322">
        - TextBindings</span></span><br><span data-ttu-id="5c919-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-324">Office 2016 on Mac</span></span><br><span data-ttu-id="5c919-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5c919-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-326">- TaskPane</span></span><br><span data-ttu-id="5c919-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-327">
        - Content</span></span></td>
    <td><span data-ttu-id="5c919-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5c919-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5c919-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5c919-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-331">- BindingEvents</span></span><br><span data-ttu-id="5c919-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-332">
        - CompressedFile</span></span><br><span data-ttu-id="5c919-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-333">
        - DocumentEvents</span></span><br><span data-ttu-id="5c919-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="5c919-334">
        - File</span></span><br><span data-ttu-id="5c919-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-335">
        - MatrixBindings</span></span><br><span data-ttu-id="5c919-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="5c919-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-337">
        - PdfFile</span></span><br><span data-ttu-id="5c919-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-338">
        - Selection</span></span><br><span data-ttu-id="5c919-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-339">
        - Settings</span></span><br><span data-ttu-id="5c919-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-340">
        - TableBindings</span></span><br><span data-ttu-id="5c919-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-341">
        - TableCoercion</span></span><br><span data-ttu-id="5c919-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-342">
        - TextBindings</span></span><br><span data-ttu-id="5c919-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5c919-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="5c919-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="5c919-345">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="5c919-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5c919-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="5c919-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5c919-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5c919-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5c919-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-350">Office on the web</span></span></td>
    <td><span data-ttu-id="5c919-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5c919-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-353">Office on Windows</span></span><br><span data-ttu-id="5c919-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5c919-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5c919-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-357">Office for Mac</span></span><br><span data-ttu-id="5c919-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5c919-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c919-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5c919-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="5c919-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="5c919-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5c919-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="5c919-362">Platform</span></span></th>
    <th><span data-ttu-id="5c919-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-363">Extension points</span></span></th>
    <th><span data-ttu-id="5c919-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="5c919-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-366">Office on the web</span></span><br><span data-ttu-id="5c919-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="5c919-367">(modern)</span></span></td>
    <td> <span data-ttu-id="5c919-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-368">- Mail Read</span></span><br><span data-ttu-id="5c919-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-369">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5c919-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="5c919-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="5c919-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-380">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-380">Office on the web</span></span><br><span data-ttu-id="5c919-381">(classique)</span><span class="sxs-lookup"><span data-stu-id="5c919-381">(classic)</span></span></td>
    <td> <span data-ttu-id="5c919-382">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-382">- Mail Read</span></span><br><span data-ttu-id="5c919-383">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-383">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5c919-391">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-392">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-392">Office on Windows</span></span><br><span data-ttu-id="5c919-393">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-394">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-394">- Mail Read</span></span><br><span data-ttu-id="5c919-395">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-395">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5c919-397">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="5c919-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5c919-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5c919-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="5c919-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="5c919-406">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-407">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-407">Office 2019 on Windows</span></span><br><span data-ttu-id="5c919-408">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-409">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-409">- Mail Read</span></span><br><span data-ttu-id="5c919-410">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-410">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5c919-412">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="5c919-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5c919-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5c919-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5c919-420">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-421">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-421">Office 2016 on Windows</span></span><br><span data-ttu-id="5c919-422">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-423">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-423">- Mail Read</span></span><br><span data-ttu-id="5c919-424">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-424">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5c919-426">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="5c919-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5c919-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5c919-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-432">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-432">Office 2013 on Windows</span></span><br><span data-ttu-id="5c919-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-434">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-434">- Mail Read</span></span><br><span data-ttu-id="5c919-435">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="5c919-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="5c919-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5c919-440">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-441">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="5c919-441">Office on iOS</span></span><br><span data-ttu-id="5c919-442">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-443">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-443">- Mail Read</span></span><br><span data-ttu-id="5c919-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5c919-450">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-451">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-451">Office on Mac</span></span><br><span data-ttu-id="5c919-452">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-453">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-453">- Mail Read</span></span><br><span data-ttu-id="5c919-454">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-454">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5c919-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5c919-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="5c919-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5c919-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="5c919-464">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-465">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-465">Office 2019 on Mac</span></span><br><span data-ttu-id="5c919-466">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-467">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-467">- Mail Read</span></span><br><span data-ttu-id="5c919-468">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-468">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5c919-476">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-477">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-477">Office 2016 on Mac</span></span><br><span data-ttu-id="5c919-478">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-479">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-479">- Mail Read</span></span><br><span data-ttu-id="5c919-480">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="5c919-480">
      - Mail Compose</span></span><br><span data-ttu-id="5c919-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5c919-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5c919-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5c919-488">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-489">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="5c919-489">Office on Android</span></span><br><span data-ttu-id="5c919-490">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-491">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="5c919-491">- Mail Read</span></span><br><span data-ttu-id="5c919-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5c919-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5c919-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5c919-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5c919-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5c919-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5c919-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5c919-498">Non disponible</span><span class="sxs-lookup"><span data-stu-id="5c919-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="5c919-499">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="5c919-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5c919-500">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="5c919-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="5c919-501">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="5c919-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="5c919-502">Word</span><span class="sxs-lookup"><span data-stu-id="5c919-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5c919-503">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="5c919-503">Platform</span></span></th>
    <th><span data-ttu-id="5c919-504">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-504">Extension points</span></span></th>
    <th><span data-ttu-id="5c919-505">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="5c919-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-507">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="5c919-508">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-508">- TaskPane</span></span><br><span data-ttu-id="5c919-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5c919-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-516">- BindingEvents</span></span><br><span data-ttu-id="5c919-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-518">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-519">
         - File</span></span><br><span data-ttu-id="5c919-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-521">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-524">
         - PdfFile</span></span><br><span data-ttu-id="5c919-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-525">
         - Selection</span></span><br><span data-ttu-id="5c919-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-526">
         - Settings</span></span><br><span data-ttu-id="5c919-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-527">
         - TableBindings</span></span><br><span data-ttu-id="5c919-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-528">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-529">
         - TextBindings</span></span><br><span data-ttu-id="5c919-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-530">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-532">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-532">Office on Windows</span></span><br><span data-ttu-id="5c919-533">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-534">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-534">- TaskPane</span></span><br><span data-ttu-id="5c919-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5c919-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-542">- BindingEvents</span></span><br><span data-ttu-id="5c919-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-543">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-545">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-546">
         - File</span></span><br><span data-ttu-id="5c919-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-548">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-551">
         - PdfFile</span></span><br><span data-ttu-id="5c919-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-552">
         - Selection</span></span><br><span data-ttu-id="5c919-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-553">
         - Settings</span></span><br><span data-ttu-id="5c919-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-554">
         - TableBindings</span></span><br><span data-ttu-id="5c919-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-555">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-556">
         - TextBindings</span></span><br><span data-ttu-id="5c919-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-557">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-559">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-559">Office 2019 on Windows</span></span><br><span data-ttu-id="5c919-560">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-561">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="5c919-561">- TaskPane</span></span><br><span data-ttu-id="5c919-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-568">- BindingEvents</span></span><br><span data-ttu-id="5c919-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-569">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-571">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-572">
         - File</span></span><br><span data-ttu-id="5c919-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-574">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-577">
         - PdfFile</span></span><br><span data-ttu-id="5c919-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-578">
         - Selection</span></span><br><span data-ttu-id="5c919-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-579">
         - Settings</span></span><br><span data-ttu-id="5c919-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-580">
         - TableBindings</span></span><br><span data-ttu-id="5c919-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-581">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-582">
         - TextBindings</span></span><br><span data-ttu-id="5c919-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-583">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-585">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-585">Office 2016 on Windows</span></span><br><span data-ttu-id="5c919-586">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-587">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5c919-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-591">- BindingEvents</span></span><br><span data-ttu-id="5c919-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-592">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-594">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-595">
         - File</span></span><br><span data-ttu-id="5c919-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-597">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-600">
         - PdfFile</span></span><br><span data-ttu-id="5c919-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-601">
         - Selection</span></span><br><span data-ttu-id="5c919-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-602">
         - Settings</span></span><br><span data-ttu-id="5c919-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-603">
         - TableBindings</span></span><br><span data-ttu-id="5c919-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-604">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-605">
         - TextBindings</span></span><br><span data-ttu-id="5c919-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-606">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-608">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-608">Office 2013 on Windows</span></span><br><span data-ttu-id="5c919-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5c919-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5c919-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-613">- BindingEvents</span></span><br><span data-ttu-id="5c919-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-614">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-616">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-617">
         - File</span></span><br><span data-ttu-id="5c919-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-619">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-622">
         - PdfFile</span></span><br><span data-ttu-id="5c919-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-623">
         - Selection</span></span><br><span data-ttu-id="5c919-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-624">
         - Settings</span></span><br><span data-ttu-id="5c919-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-625">
         - TableBindings</span></span><br><span data-ttu-id="5c919-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-626">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-627">
         - TextBindings</span></span><br><span data-ttu-id="5c919-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-628">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-630">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="5c919-630">Office on iPad</span></span><br><span data-ttu-id="5c919-631">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="5c919-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-638">- BindingEvents</span></span><br><span data-ttu-id="5c919-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-639">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-641">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-642">
         - File</span></span><br><span data-ttu-id="5c919-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-644">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-647">
         - PdfFile</span></span><br><span data-ttu-id="5c919-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-648">
         - Selection</span></span><br><span data-ttu-id="5c919-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-649">
         - Settings</span></span><br><span data-ttu-id="5c919-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-650">
         - TableBindings</span></span><br><span data-ttu-id="5c919-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-651">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-652">
         - TextBindings</span></span><br><span data-ttu-id="5c919-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-653">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-655">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-655">Office on Mac</span></span><br><span data-ttu-id="5c919-656">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-657">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-657">- TaskPane</span></span><br><span data-ttu-id="5c919-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="5c919-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-665">- BindingEvents</span></span><br><span data-ttu-id="5c919-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-666">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-668">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-669">
         - File</span></span><br><span data-ttu-id="5c919-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-671">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-674">
         - PdfFile</span></span><br><span data-ttu-id="5c919-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-675">
         - Selection</span></span><br><span data-ttu-id="5c919-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-676">
         - Settings</span></span><br><span data-ttu-id="5c919-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-677">
         - TableBindings</span></span><br><span data-ttu-id="5c919-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-678">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-679">
         - TextBindings</span></span><br><span data-ttu-id="5c919-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-680">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-682">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-682">Office 2019 on Mac</span></span><br><span data-ttu-id="5c919-683">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-684">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="5c919-684">- TaskPane</span></span><br><span data-ttu-id="5c919-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5c919-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5c919-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5c919-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="5c919-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-691">- BindingEvents</span></span><br><span data-ttu-id="5c919-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-692">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-694">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-695">
         - File</span></span><br><span data-ttu-id="5c919-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-697">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-700">
         - PdfFile</span></span><br><span data-ttu-id="5c919-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-701">
         - Selection</span></span><br><span data-ttu-id="5c919-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-702">
         - Settings</span></span><br><span data-ttu-id="5c919-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-703">
         - TableBindings</span></span><br><span data-ttu-id="5c919-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-704">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-705">
         - TextBindings</span></span><br><span data-ttu-id="5c919-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-706">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-708">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-708">Office 2016 on Mac</span></span><br><span data-ttu-id="5c919-709">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-710">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5c919-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5c919-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5c919-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-714">- BindingEvents</span></span><br><span data-ttu-id="5c919-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-715">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5c919-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="5c919-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-717">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-718">
         - File</span></span><br><span data-ttu-id="5c919-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-720">
         - MatrixBindings</span></span><br><span data-ttu-id="5c919-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="5c919-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5c919-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-723">
         - PdfFile</span></span><br><span data-ttu-id="5c919-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-724">
         - Selection</span></span><br><span data-ttu-id="5c919-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-725">
         - Settings</span></span><br><span data-ttu-id="5c919-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-726">
         - TableBindings</span></span><br><span data-ttu-id="5c919-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-727">
         - TableCoercion</span></span><br><span data-ttu-id="5c919-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5c919-728">
         - TextBindings</span></span><br><span data-ttu-id="5c919-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-729">
         - TextCoercion</span></span><br><span data-ttu-id="5c919-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5c919-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="5c919-731">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="5c919-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5c919-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5c919-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5c919-733">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="5c919-733">Platform</span></span></th>
    <th><span data-ttu-id="5c919-734">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-734">Extension points</span></span></th>
    <th><span data-ttu-id="5c919-735">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="5c919-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-737">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="5c919-738">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-738">- Content</span></span><br><span data-ttu-id="5c919-739">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-739">
         - TaskPane</span></span><br><span data-ttu-id="5c919-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5c919-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5c919-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-745">- ActiveView</span></span><br><span data-ttu-id="5c919-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-746">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-747">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-748">
         - File</span></span><br><span data-ttu-id="5c919-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-749">
         - PdfFile</span></span><br><span data-ttu-id="5c919-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-750">
         - Selection</span></span><br><span data-ttu-id="5c919-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-751">
         - Settings</span></span><br><span data-ttu-id="5c919-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-753">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-753">Office on Windows</span></span><br><span data-ttu-id="5c919-754">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-755">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-755">- Content</span></span><br><span data-ttu-id="5c919-756">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-756">
         - TaskPane</span></span><br><span data-ttu-id="5c919-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5c919-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5c919-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-762">- ActiveView</span></span><br><span data-ttu-id="5c919-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-763">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-764">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-765">
         - File</span></span><br><span data-ttu-id="5c919-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-766">
         - PdfFile</span></span><br><span data-ttu-id="5c919-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-767">
         - Selection</span></span><br><span data-ttu-id="5c919-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-768">
         - Settings</span></span><br><span data-ttu-id="5c919-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-770">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-770">Office 2019 on Windows</span></span><br><span data-ttu-id="5c919-771">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-772">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-772">- Content</span></span><br><span data-ttu-id="5c919-773">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-773">
         - TaskPane</span></span><br><span data-ttu-id="5c919-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-777">- ActiveView</span></span><br><span data-ttu-id="5c919-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-778">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-779">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-780">
         - File</span></span><br><span data-ttu-id="5c919-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-781">
         - PdfFile</span></span><br><span data-ttu-id="5c919-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-782">
         - Selection</span></span><br><span data-ttu-id="5c919-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-783">
         - Settings</span></span><br><span data-ttu-id="5c919-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-785">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-785">Office 2016 on Windows</span></span><br><span data-ttu-id="5c919-786">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-787">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-787">- Content</span></span><br><span data-ttu-id="5c919-788">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5c919-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5c919-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-791">- ActiveView</span></span><br><span data-ttu-id="5c919-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-792">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-793">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-794">
         - File</span></span><br><span data-ttu-id="5c919-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-795">
         - PdfFile</span></span><br><span data-ttu-id="5c919-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-796">
         - Selection</span></span><br><span data-ttu-id="5c919-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-797">
         - Settings</span></span><br><span data-ttu-id="5c919-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-799">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-799">Office 2013 on Windows</span></span><br><span data-ttu-id="5c919-800">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-801">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-801">- Content</span></span><br><span data-ttu-id="5c919-802">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="5c919-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5c919-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5c919-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-805">- ActiveView</span></span><br><span data-ttu-id="5c919-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-806">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-807">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-808">
         - File</span></span><br><span data-ttu-id="5c919-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-809">
         - PdfFile</span></span><br><span data-ttu-id="5c919-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-810">
         - Selection</span></span><br><span data-ttu-id="5c919-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-811">
         - Settings</span></span><br><span data-ttu-id="5c919-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-813">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="5c919-813">Office on iPad</span></span><br><span data-ttu-id="5c919-814">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-815">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-815">- Content</span></span><br><span data-ttu-id="5c919-816">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5c919-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-820">- ActiveView</span></span><br><span data-ttu-id="5c919-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-821">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-822">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-823">
         - File</span></span><br><span data-ttu-id="5c919-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-824">
         - PdfFile</span></span><br><span data-ttu-id="5c919-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-825">
         - Selection</span></span><br><span data-ttu-id="5c919-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-826">
         - Settings</span></span><br><span data-ttu-id="5c919-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-828">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-828">Office on Mac</span></span><br><span data-ttu-id="5c919-829">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="5c919-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5c919-830">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-830">- Content</span></span><br><span data-ttu-id="5c919-831">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-831">
         - TaskPane</span></span><br><span data-ttu-id="5c919-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5c919-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5c919-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5c919-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5c919-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-837">- ActiveView</span></span><br><span data-ttu-id="5c919-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-838">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-839">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-840">
         - File</span></span><br><span data-ttu-id="5c919-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-841">
         - PdfFile</span></span><br><span data-ttu-id="5c919-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-842">
         - Selection</span></span><br><span data-ttu-id="5c919-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-843">
         - Settings</span></span><br><span data-ttu-id="5c919-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-845">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-845">Office 2019 on Mac</span></span><br><span data-ttu-id="5c919-846">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-847">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-847">- Content</span></span><br><span data-ttu-id="5c919-848">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-848">
         - TaskPane</span></span><br><span data-ttu-id="5c919-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-852">- ActiveView</span></span><br><span data-ttu-id="5c919-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-853">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-854">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-855">
         - File</span></span><br><span data-ttu-id="5c919-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-856">
         - PdfFile</span></span><br><span data-ttu-id="5c919-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-857">
         - Selection</span></span><br><span data-ttu-id="5c919-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-858">
         - Settings</span></span><br><span data-ttu-id="5c919-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-860">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-860">Office 2016 on Mac</span></span><br><span data-ttu-id="5c919-861">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-862">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-862">- Content</span></span><br><span data-ttu-id="5c919-863">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5c919-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5c919-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5c919-866">- ActiveView</span></span><br><span data-ttu-id="5c919-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5c919-867">
         - CompressedFile</span></span><br><span data-ttu-id="5c919-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-868">
         - DocumentEvents</span></span><br><span data-ttu-id="5c919-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="5c919-869">
         - File</span></span><br><span data-ttu-id="5c919-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5c919-870">
         - PdfFile</span></span><br><span data-ttu-id="5c919-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-871">
         - Selection</span></span><br><span data-ttu-id="5c919-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-872">
         - Settings</span></span><br><span data-ttu-id="5c919-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5c919-874">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="5c919-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="5c919-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="5c919-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5c919-876">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="5c919-876">Platform</span></span></th>
    <th><span data-ttu-id="5c919-877">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-877">Extension points</span></span></th>
    <th><span data-ttu-id="5c919-878">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="5c919-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-880">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="5c919-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="5c919-881">- Contenu</span><span class="sxs-lookup"><span data-stu-id="5c919-881">- Content</span></span><br><span data-ttu-id="5c919-882">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-882">
         - TaskPane</span></span><br><span data-ttu-id="5c919-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="5c919-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5c919-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5c919-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5c919-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5c919-887">- DocumentEvents</span></span><br><span data-ttu-id="5c919-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="5c919-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5c919-889">
         - Settings</span></span><br><span data-ttu-id="5c919-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="5c919-891">Projet</span><span class="sxs-lookup"><span data-stu-id="5c919-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5c919-892">Plateforme</span><span class="sxs-lookup"><span data-stu-id="5c919-892">Platform</span></span></th>
    <th><span data-ttu-id="5c919-893">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="5c919-893">Extension points</span></span></th>
    <th><span data-ttu-id="5c919-894">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="5c919-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="5c919-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-896">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-896">Office 2019 on Windows</span></span><br><span data-ttu-id="5c919-897">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-898">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-900">- Selection</span></span><br><span data-ttu-id="5c919-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-902">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-902">Office 2016 on Windows</span></span><br><span data-ttu-id="5c919-903">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-904">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-906">- Selection</span></span><br><span data-ttu-id="5c919-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5c919-908">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="5c919-908">Office 2013 on Windows</span></span><br><span data-ttu-id="5c919-909">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="5c919-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5c919-910">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="5c919-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5c919-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5c919-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5c919-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="5c919-912">- Selection</span></span><br><span data-ttu-id="5c919-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5c919-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5c919-914">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5c919-914">See also</span></span>

- [<span data-ttu-id="5c919-915">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="5c919-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5c919-916">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="5c919-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5c919-917">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="5c919-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="5c919-918">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="5c919-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="5c919-919">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="5c919-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="5c919-920">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="5c919-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="5c919-921">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="5c919-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="5c919-922">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="5c919-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="5c919-923">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="5c919-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="5c919-924">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="5c919-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="5c919-925">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="5c919-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="5c919-926">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="5c919-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)