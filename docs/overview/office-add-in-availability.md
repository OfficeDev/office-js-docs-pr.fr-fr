---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901947"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7e918-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="7e918-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7e918-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="7e918-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7e918-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="7e918-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="7e918-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="7e918-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="7e918-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7e918-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7e918-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7e918-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7e918-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7e918-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7e918-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="7e918-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-114">- TaskPane</span></span><br><span data-ttu-id="7e918-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-115">
        - Content</span></span><br><span data-ttu-id="7e918-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-116">
        - Custom Functions</span></span><br><span data-ttu-id="7e918-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="7e918-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7e918-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7e918-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7e918-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-128">
        - BindingEvents</span></span><br><span data-ttu-id="7e918-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-129">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-130">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-131">
        - File</span></span><br><span data-ttu-id="7e918-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-132">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-134">
        - Selection</span></span><br><span data-ttu-id="7e918-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-135">
        - Settings</span></span><br><span data-ttu-id="7e918-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-136">
        - TableBindings</span></span><br><span data-ttu-id="7e918-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-137">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-138">
        - TextBindings</span></span><br><span data-ttu-id="7e918-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-140">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-140">Office on Windows</span></span><br><span data-ttu-id="7e918-141">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-142">- TaskPane</span></span><br><span data-ttu-id="7e918-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-143">
        - Content</span></span><br><span data-ttu-id="7e918-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-144">
        - Custom Functions</span></span><br><span data-ttu-id="7e918-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="7e918-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7e918-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7e918-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7e918-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7e918-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-158">
        - BindingEvents</span></span><br><span data-ttu-id="7e918-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-159">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-160">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-161">
        - File</span></span><br><span data-ttu-id="7e918-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-162">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-164">
        - Selection</span></span><br><span data-ttu-id="7e918-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-165">
        - Settings</span></span><br><span data-ttu-id="7e918-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-166">
        - TableBindings</span></span><br><span data-ttu-id="7e918-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-167">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-168">
        - TextBindings</span></span><br><span data-ttu-id="7e918-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-170">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-170">Office 2019 on Windows</span></span><br><span data-ttu-id="7e918-171">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7e918-172">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-172">- TaskPane</span></span><br><span data-ttu-id="7e918-173">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-173">
        - Content</span></span><br><span data-ttu-id="7e918-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e918-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-185">- BindingEvents</span></span><br><span data-ttu-id="7e918-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-186">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-187">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-188">
        - File</span></span><br><span data-ttu-id="7e918-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-189">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-191">
        - Selection</span></span><br><span data-ttu-id="7e918-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-192">
        - Settings</span></span><br><span data-ttu-id="7e918-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-193">
        - TableBindings</span></span><br><span data-ttu-id="7e918-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-194">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-195">
        - TextBindings</span></span><br><span data-ttu-id="7e918-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-197">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-197">Office 2016 on Windows</span></span><br><span data-ttu-id="7e918-198">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7e918-199">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-199">- TaskPane</span></span><br><span data-ttu-id="7e918-200">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-200">
        - Content</span></span></td>
    <td><span data-ttu-id="7e918-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7e918-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-204">- BindingEvents</span></span><br><span data-ttu-id="7e918-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-205">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-206">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-207">
        - File</span></span><br><span data-ttu-id="7e918-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-208">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-210">
        - Selection</span></span><br><span data-ttu-id="7e918-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-211">
        - Settings</span></span><br><span data-ttu-id="7e918-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-212">
        - TableBindings</span></span><br><span data-ttu-id="7e918-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-213">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-214">
        - TextBindings</span></span><br><span data-ttu-id="7e918-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-216">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-216">Office 2013 on Windows</span></span><br><span data-ttu-id="7e918-217">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7e918-218">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-218">
        - TaskPane</span></span><br><span data-ttu-id="7e918-219">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7e918-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e918-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7e918-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-222">
        - BindingEvents</span></span><br><span data-ttu-id="7e918-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-223">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-224">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-225">
        - File</span></span><br><span data-ttu-id="7e918-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-226">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-228">
        - Selection</span></span><br><span data-ttu-id="7e918-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-229">
        - Settings</span></span><br><span data-ttu-id="7e918-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-230">
        - TableBindings</span></span><br><span data-ttu-id="7e918-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-231">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-232">
        - TextBindings</span></span><br><span data-ttu-id="7e918-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-234">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="7e918-234">Office on iPad</span></span><br><span data-ttu-id="7e918-235">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7e918-236">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-236">- TaskPane</span></span><br><span data-ttu-id="7e918-237">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-237">
        - Content</span></span></td>
    <td><span data-ttu-id="7e918-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7e918-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7e918-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-249">- BindingEvents</span></span><br><span data-ttu-id="7e918-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-250">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-251">
        - File</span></span><br><span data-ttu-id="7e918-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-252">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-254">
        - Selection</span></span><br><span data-ttu-id="7e918-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-255">
        - Settings</span></span><br><span data-ttu-id="7e918-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-256">
        - TableBindings</span></span><br><span data-ttu-id="7e918-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-257">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-258">
        - TextBindings</span></span><br><span data-ttu-id="7e918-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-260">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-260">Office on Mac</span></span><br><span data-ttu-id="7e918-261">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7e918-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-262">- TaskPane</span></span><br><span data-ttu-id="7e918-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-263">
        - Content</span></span><br><span data-ttu-id="7e918-264">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-264">
        - Custom Functions</span></span><br><span data-ttu-id="7e918-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e918-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7e918-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7e918-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7e918-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-278">- BindingEvents</span></span><br><span data-ttu-id="7e918-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-279">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-280">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-281">
        - File</span></span><br><span data-ttu-id="7e918-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-282">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-284">
        - PdfFile</span></span><br><span data-ttu-id="7e918-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-285">
        - Selection</span></span><br><span data-ttu-id="7e918-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-286">
        - Settings</span></span><br><span data-ttu-id="7e918-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-287">
        - TableBindings</span></span><br><span data-ttu-id="7e918-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-288">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-289">
        - TextBindings</span></span><br><span data-ttu-id="7e918-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-291">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-291">Office 2019 on Mac</span></span><br><span data-ttu-id="7e918-292">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7e918-293">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-293">- TaskPane</span></span><br><span data-ttu-id="7e918-294">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-294">
        - Content</span></span><br><span data-ttu-id="7e918-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e918-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e918-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e918-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e918-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e918-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e918-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e918-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e918-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-306">- BindingEvents</span></span><br><span data-ttu-id="7e918-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-307">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-308">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-309">
        - File</span></span><br><span data-ttu-id="7e918-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-310">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-312">
        - PdfFile</span></span><br><span data-ttu-id="7e918-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-313">
        - Selection</span></span><br><span data-ttu-id="7e918-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-314">
        - Settings</span></span><br><span data-ttu-id="7e918-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-315">
        - TableBindings</span></span><br><span data-ttu-id="7e918-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-316">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-317">
        - TextBindings</span></span><br><span data-ttu-id="7e918-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-319">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-319">Office 2016 on Mac</span></span><br><span data-ttu-id="7e918-320">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7e918-321">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-321">- TaskPane</span></span><br><span data-ttu-id="7e918-322">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-322">
        - Content</span></span></td>
    <td><span data-ttu-id="7e918-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e918-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7e918-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7e918-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-326">- BindingEvents</span></span><br><span data-ttu-id="7e918-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-327">
        - CompressedFile</span></span><br><span data-ttu-id="7e918-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-328">
        - DocumentEvents</span></span><br><span data-ttu-id="7e918-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e918-329">
        - File</span></span><br><span data-ttu-id="7e918-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-330">
        - MatrixBindings</span></span><br><span data-ttu-id="7e918-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e918-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-332">
        - PdfFile</span></span><br><span data-ttu-id="7e918-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-333">
        - Selection</span></span><br><span data-ttu-id="7e918-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-334">
        - Settings</span></span><br><span data-ttu-id="7e918-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-335">
        - TableBindings</span></span><br><span data-ttu-id="7e918-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-336">
        - TableCoercion</span></span><br><span data-ttu-id="7e918-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-337">
        - TextBindings</span></span><br><span data-ttu-id="7e918-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7e918-339">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7e918-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="7e918-340">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7e918-341">Plateforme</span><span class="sxs-lookup"><span data-stu-id="7e918-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7e918-342">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7e918-343">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7e918-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-345">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-345">Office on the web</span></span></td>
    <td><span data-ttu-id="7e918-346">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7e918-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-348">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-348">Office on Windows</span></span><br><span data-ttu-id="7e918-349">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7e918-350">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7e918-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-352">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-352">Office for Mac</span></span><br><span data-ttu-id="7e918-353">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7e918-354">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7e918-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7e918-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="7e918-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="7e918-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e918-357">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7e918-357">Platform</span></span></th>
    <th><span data-ttu-id="7e918-358">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-358">Extension points</span></span></th>
    <th><span data-ttu-id="7e918-359">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e918-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-361">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-361">Office on the web</span></span><br><span data-ttu-id="7e918-362">(moderne)</span><span class="sxs-lookup"><span data-stu-id="7e918-362">(modern)</span></span></td>
    <td> <span data-ttu-id="7e918-363">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-363">- Mail Read</span></span><br><span data-ttu-id="7e918-364">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-364">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e918-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7e918-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7e918-374">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-375">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-375">Office on the web</span></span><br><span data-ttu-id="7e918-376">(classique)</span><span class="sxs-lookup"><span data-stu-id="7e918-376">(classic)</span></span></td>
    <td> <span data-ttu-id="7e918-377">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-377">- Mail Read</span></span><br><span data-ttu-id="7e918-378">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-378">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e918-386">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-387">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-387">Office on Windows</span></span><br><span data-ttu-id="7e918-388">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-389">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-389">- Mail Read</span></span><br><span data-ttu-id="7e918-390">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-390">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e918-392">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7e918-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e918-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e918-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7e918-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7e918-401">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-401">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-402">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-402">Office 2019 on Windows</span></span><br><span data-ttu-id="7e918-403">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-403">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-404">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-404">- Mail Read</span></span><br><span data-ttu-id="7e918-405">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-405">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e918-407">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7e918-407">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e918-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e918-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7e918-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-416">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-416">Office 2016 on Windows</span></span><br><span data-ttu-id="7e918-417">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-418">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-418">- Mail Read</span></span><br><span data-ttu-id="7e918-419">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-419">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e918-421">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="7e918-421">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e918-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7e918-426">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-427">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-427">Office 2013 on Windows</span></span><br><span data-ttu-id="7e918-428">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-428">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-429">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-429">- Mail Read</span></span><br><span data-ttu-id="7e918-430">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-430">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="7e918-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="7e918-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7e918-435">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-436">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="7e918-436">Office on iOS</span></span><br><span data-ttu-id="7e918-437">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-437">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-438">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-438">- Mail Read</span></span><br><span data-ttu-id="7e918-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7e918-445">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-446">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-446">Office on Mac</span></span><br><span data-ttu-id="7e918-447">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-447">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-448">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-448">- Mail Read</span></span><br><span data-ttu-id="7e918-449">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-449">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e918-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e918-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7e918-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e918-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7e918-459">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-460">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-460">Office 2019 on Mac</span></span><br><span data-ttu-id="7e918-461">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-462">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-462">- Mail Read</span></span><br><span data-ttu-id="7e918-463">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-463">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e918-471">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-472">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-472">Office 2016 on Mac</span></span><br><span data-ttu-id="7e918-473">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-474">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-474">- Mail Read</span></span><br><span data-ttu-id="7e918-475">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="7e918-475">
      - Mail Compose</span></span><br><span data-ttu-id="7e918-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e918-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e918-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e918-483">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-484">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="7e918-484">Office on Android</span></span><br><span data-ttu-id="7e918-485">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-486">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="7e918-486">- Mail Read</span></span><br><span data-ttu-id="7e918-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e918-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e918-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e918-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e918-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e918-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e918-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7e918-493">Non disponible</span><span class="sxs-lookup"><span data-stu-id="7e918-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="7e918-494">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7e918-494">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7e918-495">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="7e918-495">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="7e918-496">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="7e918-496">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="7e918-497">Word</span><span class="sxs-lookup"><span data-stu-id="7e918-497">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e918-498">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7e918-498">Platform</span></span></th>
    <th><span data-ttu-id="7e918-499">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-499">Extension points</span></span></th>
    <th><span data-ttu-id="7e918-500">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-500">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e918-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-502">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-502">Office on the web</span></span></td>
    <td> <span data-ttu-id="7e918-503">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-503">- TaskPane</span></span><br><span data-ttu-id="7e918-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7e918-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-511">- BindingEvents</span></span><br><span data-ttu-id="7e918-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-513">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-514">
         - File</span></span><br><span data-ttu-id="7e918-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-516">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-519">
         - PdfFile</span></span><br><span data-ttu-id="7e918-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-520">
         - Selection</span></span><br><span data-ttu-id="7e918-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-521">
         - Settings</span></span><br><span data-ttu-id="7e918-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-522">
         - TableBindings</span></span><br><span data-ttu-id="7e918-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-523">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-524">
         - TextBindings</span></span><br><span data-ttu-id="7e918-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-525">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-526">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-527">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-527">Office on Windows</span></span><br><span data-ttu-id="7e918-528">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-528">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-529">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-529">- TaskPane</span></span><br><span data-ttu-id="7e918-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7e918-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-537">- BindingEvents</span></span><br><span data-ttu-id="7e918-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-538">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-540">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-541">
         - File</span></span><br><span data-ttu-id="7e918-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-543">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-546">
         - PdfFile</span></span><br><span data-ttu-id="7e918-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-547">
         - Selection</span></span><br><span data-ttu-id="7e918-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-548">
         - Settings</span></span><br><span data-ttu-id="7e918-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-549">
         - TableBindings</span></span><br><span data-ttu-id="7e918-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-550">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-551">
         - TextBindings</span></span><br><span data-ttu-id="7e918-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-552">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-553">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-554">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-554">Office 2019 on Windows</span></span><br><span data-ttu-id="7e918-555">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-555">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-556">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="7e918-556">- TaskPane</span></span><br><span data-ttu-id="7e918-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-563">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-563">- BindingEvents</span></span><br><span data-ttu-id="7e918-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-564">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-565">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-565">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-566">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-567">
         - File</span></span><br><span data-ttu-id="7e918-568">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-568">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-569">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-569">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-570">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-570">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-571">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-571">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-572">
         - PdfFile</span></span><br><span data-ttu-id="7e918-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-573">
         - Selection</span></span><br><span data-ttu-id="7e918-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-574">
         - Settings</span></span><br><span data-ttu-id="7e918-575">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-575">
         - TableBindings</span></span><br><span data-ttu-id="7e918-576">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-576">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-577">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-577">
         - TextBindings</span></span><br><span data-ttu-id="7e918-578">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-578">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-579">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-579">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-580">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-580">Office 2016 on Windows</span></span><br><span data-ttu-id="7e918-581">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-581">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-582">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-582">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7e918-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-586">- BindingEvents</span></span><br><span data-ttu-id="7e918-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-587">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-589">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-590">
         - File</span></span><br><span data-ttu-id="7e918-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-592">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-595">
         - PdfFile</span></span><br><span data-ttu-id="7e918-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-596">
         - Selection</span></span><br><span data-ttu-id="7e918-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-597">
         - Settings</span></span><br><span data-ttu-id="7e918-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-598">
         - TableBindings</span></span><br><span data-ttu-id="7e918-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-599">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-600">
         - TextBindings</span></span><br><span data-ttu-id="7e918-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-601">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-603">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-603">Office 2013 on Windows</span></span><br><span data-ttu-id="7e918-604">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-605">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e918-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7e918-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-608">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-608">- BindingEvents</span></span><br><span data-ttu-id="7e918-609">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-609">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-610">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-610">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-611">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-611">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-612">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-612">
         - File</span></span><br><span data-ttu-id="7e918-613">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-613">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-614">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-614">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-615">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-615">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-616">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-616">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-617">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-617">
         - PdfFile</span></span><br><span data-ttu-id="7e918-618">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-618">
         - Selection</span></span><br><span data-ttu-id="7e918-619">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-619">
         - Settings</span></span><br><span data-ttu-id="7e918-620">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-620">
         - TableBindings</span></span><br><span data-ttu-id="7e918-621">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-621">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-622">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-622">
         - TextBindings</span></span><br><span data-ttu-id="7e918-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-623">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-624">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-624">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-625">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="7e918-625">Office on iPad</span></span><br><span data-ttu-id="7e918-626">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-626">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-627">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-627">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7e918-633">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-633">- BindingEvents</span></span><br><span data-ttu-id="7e918-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-634">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-635">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-635">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-636">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-636">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-637">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-637">
         - File</span></span><br><span data-ttu-id="7e918-638">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-638">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-639">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-639">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-640">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-640">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-641">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-641">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-642">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-642">
         - PdfFile</span></span><br><span data-ttu-id="7e918-643">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-643">
         - Selection</span></span><br><span data-ttu-id="7e918-644">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-644">
         - Settings</span></span><br><span data-ttu-id="7e918-645">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-645">
         - TableBindings</span></span><br><span data-ttu-id="7e918-646">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-646">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-647">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-647">
         - TextBindings</span></span><br><span data-ttu-id="7e918-648">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-648">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-649">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-649">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-650">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-650">Office on Mac</span></span><br><span data-ttu-id="7e918-651">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-651">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-652">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-652">- TaskPane</span></span><br><span data-ttu-id="7e918-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="7e918-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-660">- BindingEvents</span></span><br><span data-ttu-id="7e918-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-661">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-663">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-664">
         - File</span></span><br><span data-ttu-id="7e918-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-666">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-669">
         - PdfFile</span></span><br><span data-ttu-id="7e918-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-670">
         - Selection</span></span><br><span data-ttu-id="7e918-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-671">
         - Settings</span></span><br><span data-ttu-id="7e918-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-672">
         - TableBindings</span></span><br><span data-ttu-id="7e918-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-673">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-674">
         - TextBindings</span></span><br><span data-ttu-id="7e918-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-675">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-677">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-677">Office 2019 on Mac</span></span><br><span data-ttu-id="7e918-678">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-678">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-679">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="7e918-679">- TaskPane</span></span><br><span data-ttu-id="7e918-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7e918-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e918-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7e918-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7e918-686">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-686">- BindingEvents</span></span><br><span data-ttu-id="7e918-687">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-687">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-688">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-688">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-689">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-689">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-690">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-690">
         - File</span></span><br><span data-ttu-id="7e918-691">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-691">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-692">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-692">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-693">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-693">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-694">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-694">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-695">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-695">
         - PdfFile</span></span><br><span data-ttu-id="7e918-696">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-696">
         - Selection</span></span><br><span data-ttu-id="7e918-697">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-697">
         - Settings</span></span><br><span data-ttu-id="7e918-698">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-698">
         - TableBindings</span></span><br><span data-ttu-id="7e918-699">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-699">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-700">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-700">
         - TextBindings</span></span><br><span data-ttu-id="7e918-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-701">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-702">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-702">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-703">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-703">Office 2016 on Mac</span></span><br><span data-ttu-id="7e918-704">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-704">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-705">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-705">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7e918-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e918-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7e918-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-709">- BindingEvents</span></span><br><span data-ttu-id="7e918-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-710">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e918-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e918-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-712">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-713">
         - File</span></span><br><span data-ttu-id="7e918-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-715">
         - MatrixBindings</span></span><br><span data-ttu-id="7e918-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e918-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e918-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-718">
         - PdfFile</span></span><br><span data-ttu-id="7e918-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-719">
         - Selection</span></span><br><span data-ttu-id="7e918-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-720">
         - Settings</span></span><br><span data-ttu-id="7e918-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-721">
         - TableBindings</span></span><br><span data-ttu-id="7e918-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-722">
         - TableCoercion</span></span><br><span data-ttu-id="7e918-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e918-723">
         - TextBindings</span></span><br><span data-ttu-id="7e918-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-724">
         - TextCoercion</span></span><br><span data-ttu-id="7e918-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e918-725">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="7e918-726">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7e918-726">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7e918-727">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7e918-727">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e918-728">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7e918-728">Platform</span></span></th>
    <th><span data-ttu-id="7e918-729">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-729">Extension points</span></span></th>
    <th><span data-ttu-id="7e918-730">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-730">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e918-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-732">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-732">Office on the web</span></span></td>
    <td> <span data-ttu-id="7e918-733">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-733">- Content</span></span><br><span data-ttu-id="7e918-734">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-734">
         - TaskPane</span></span><br><span data-ttu-id="7e918-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7e918-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7e918-740">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-740">- ActiveView</span></span><br><span data-ttu-id="7e918-741">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-741">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-742">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-742">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-743">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-743">
         - File</span></span><br><span data-ttu-id="7e918-744">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-744">
         - PdfFile</span></span><br><span data-ttu-id="7e918-745">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-745">
         - Selection</span></span><br><span data-ttu-id="7e918-746">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-746">
         - Settings</span></span><br><span data-ttu-id="7e918-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-747">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-748">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-748">Office on Windows</span></span><br><span data-ttu-id="7e918-749">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-749">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-750">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-750">- Content</span></span><br><span data-ttu-id="7e918-751">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-751">
         - TaskPane</span></span><br><span data-ttu-id="7e918-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7e918-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7e918-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-757">- ActiveView</span></span><br><span data-ttu-id="7e918-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-758">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-759">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-760">
         - File</span></span><br><span data-ttu-id="7e918-761">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-761">
         - PdfFile</span></span><br><span data-ttu-id="7e918-762">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-762">
         - Selection</span></span><br><span data-ttu-id="7e918-763">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-763">
         - Settings</span></span><br><span data-ttu-id="7e918-764">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-764">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-765">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-765">Office 2019 on Windows</span></span><br><span data-ttu-id="7e918-766">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-766">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-767">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-767">- Content</span></span><br><span data-ttu-id="7e918-768">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-768">
         - TaskPane</span></span><br><span data-ttu-id="7e918-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-772">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-772">- ActiveView</span></span><br><span data-ttu-id="7e918-773">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-773">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-774">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-774">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-775">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-775">
         - File</span></span><br><span data-ttu-id="7e918-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-776">
         - PdfFile</span></span><br><span data-ttu-id="7e918-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-777">
         - Selection</span></span><br><span data-ttu-id="7e918-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-778">
         - Settings</span></span><br><span data-ttu-id="7e918-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-780">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-780">Office 2016 on Windows</span></span><br><span data-ttu-id="7e918-781">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-782">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-782">- Content</span></span><br><span data-ttu-id="7e918-783">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-783">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e918-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7e918-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-786">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-786">- ActiveView</span></span><br><span data-ttu-id="7e918-787">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-787">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-788">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-788">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-789">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-789">
         - File</span></span><br><span data-ttu-id="7e918-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-790">
         - PdfFile</span></span><br><span data-ttu-id="7e918-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-791">
         - Selection</span></span><br><span data-ttu-id="7e918-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-792">
         - Settings</span></span><br><span data-ttu-id="7e918-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-794">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-794">Office 2013 on Windows</span></span><br><span data-ttu-id="7e918-795">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-795">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-796">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-796">- Content</span></span><br><span data-ttu-id="7e918-797">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-797">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7e918-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e918-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7e918-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-800">- ActiveView</span></span><br><span data-ttu-id="7e918-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-801">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-802">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-803">
         - File</span></span><br><span data-ttu-id="7e918-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-804">
         - PdfFile</span></span><br><span data-ttu-id="7e918-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-805">
         - Selection</span></span><br><span data-ttu-id="7e918-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-806">
         - Settings</span></span><br><span data-ttu-id="7e918-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-808">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="7e918-808">Office on iPad</span></span><br><span data-ttu-id="7e918-809">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-810">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-810">- Content</span></span><br><span data-ttu-id="7e918-811">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7e918-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-815">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-815">- ActiveView</span></span><br><span data-ttu-id="7e918-816">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-816">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-817">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-817">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-818">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-818">
         - File</span></span><br><span data-ttu-id="7e918-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-819">
         - PdfFile</span></span><br><span data-ttu-id="7e918-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-820">
         - Selection</span></span><br><span data-ttu-id="7e918-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-821">
         - Settings</span></span><br><span data-ttu-id="7e918-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-823">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-823">Office on Mac</span></span><br><span data-ttu-id="7e918-824">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="7e918-824">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7e918-825">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-825">- Content</span></span><br><span data-ttu-id="7e918-826">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-826">
         - TaskPane</span></span><br><span data-ttu-id="7e918-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7e918-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7e918-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e918-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7e918-832">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-832">- ActiveView</span></span><br><span data-ttu-id="7e918-833">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-833">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-834">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-834">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-835">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-835">
         - File</span></span><br><span data-ttu-id="7e918-836">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-836">
         - PdfFile</span></span><br><span data-ttu-id="7e918-837">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-837">
         - Selection</span></span><br><span data-ttu-id="7e918-838">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-838">
         - Settings</span></span><br><span data-ttu-id="7e918-839">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-839">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-840">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-840">Office 2019 on Mac</span></span><br><span data-ttu-id="7e918-841">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-841">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-842">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-842">- Content</span></span><br><span data-ttu-id="7e918-843">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-843">
         - TaskPane</span></span><br><span data-ttu-id="7e918-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-847">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-847">- ActiveView</span></span><br><span data-ttu-id="7e918-848">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-848">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-849">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-849">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-850">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-850">
         - File</span></span><br><span data-ttu-id="7e918-851">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-851">
         - PdfFile</span></span><br><span data-ttu-id="7e918-852">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-852">
         - Selection</span></span><br><span data-ttu-id="7e918-853">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-853">
         - Settings</span></span><br><span data-ttu-id="7e918-854">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-854">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-855">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-855">Office 2016 on Mac</span></span><br><span data-ttu-id="7e918-856">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-856">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-857">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-857">- Content</span></span><br><span data-ttu-id="7e918-858">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-858">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e918-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7e918-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-861">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e918-861">- ActiveView</span></span><br><span data-ttu-id="7e918-862">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e918-862">
         - CompressedFile</span></span><br><span data-ttu-id="7e918-863">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-863">
         - DocumentEvents</span></span><br><span data-ttu-id="7e918-864">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e918-864">
         - File</span></span><br><span data-ttu-id="7e918-865">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e918-865">
         - PdfFile</span></span><br><span data-ttu-id="7e918-866">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-866">
         - Selection</span></span><br><span data-ttu-id="7e918-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-867">
         - Settings</span></span><br><span data-ttu-id="7e918-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7e918-869">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="7e918-869">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="7e918-870">OneNote</span><span class="sxs-lookup"><span data-stu-id="7e918-870">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e918-871">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="7e918-871">Platform</span></span></th>
    <th><span data-ttu-id="7e918-872">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-872">Extension points</span></span></th>
    <th><span data-ttu-id="7e918-873">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-873">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e918-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-875">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7e918-875">Office on the web</span></span></td>
    <td> <span data-ttu-id="7e918-876">- Contenu</span><span class="sxs-lookup"><span data-stu-id="7e918-876">- Content</span></span><br><span data-ttu-id="7e918-877">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-877">
         - TaskPane</span></span><br><span data-ttu-id="7e918-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="7e918-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e918-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7e918-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7e918-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-882">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e918-882">- DocumentEvents</span></span><br><span data-ttu-id="7e918-883">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-883">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e918-884">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e918-884">
         - Settings</span></span><br><span data-ttu-id="7e918-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-885">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7e918-886">Projet</span><span class="sxs-lookup"><span data-stu-id="7e918-886">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e918-887">Plateforme</span><span class="sxs-lookup"><span data-stu-id="7e918-887">Platform</span></span></th>
    <th><span data-ttu-id="7e918-888">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="7e918-888">Extension points</span></span></th>
    <th><span data-ttu-id="7e918-889">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="7e918-889">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e918-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="7e918-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-891">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-891">Office 2019 on Windows</span></span><br><span data-ttu-id="7e918-892">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-893">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-895">- Selection</span></span><br><span data-ttu-id="7e918-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-897">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-897">Office 2016 on Windows</span></span><br><span data-ttu-id="7e918-898">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-899">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-901">- Selection</span></span><br><span data-ttu-id="7e918-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-902">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e918-903">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="7e918-903">Office 2013 on Windows</span></span><br><span data-ttu-id="7e918-904">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="7e918-904">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7e918-905">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="7e918-905">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e918-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e918-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e918-907">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e918-907">- Selection</span></span><br><span data-ttu-id="7e918-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e918-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7e918-909">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7e918-909">See also</span></span>

- [<span data-ttu-id="7e918-910">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="7e918-910">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7e918-911">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7e918-911">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="7e918-912">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="7e918-912">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="7e918-913">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="7e918-913">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="7e918-914">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="7e918-914">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="7e918-915">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="7e918-915">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="7e918-916">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="7e918-916">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="7e918-917">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="7e918-917">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="7e918-918">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="7e918-918">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="7e918-919">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="7e918-919">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="7e918-920">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="7e918-920">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
