---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 04/07/2020
localization_priority: Priority
ms.openlocfilehash: 823fd53e71c71f4a845f9a7b5c6177ad3f14745f
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185616"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b35e4-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="b35e4-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b35e4-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="b35e4-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b35e4-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="b35e4-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b35e4-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b35e4-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b35e4-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b35e4-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b35e4-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b35e4-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b35e4-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b35e4-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b35e4-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b35e4-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-114">- TaskPane</span></span><br><span data-ttu-id="b35e4-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-115">
        - Content</span></span><br><span data-ttu-id="b35e4-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-116">
        - Custom Functions</span></span><br><span data-ttu-id="b35e4-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b35e4-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b35e4-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b35e4-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b35e4-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="b35e4-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-130">
        - BindingEvents</span></span><br><span data-ttu-id="b35e4-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-131">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-132">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-133">
        - File</span></span><br><span data-ttu-id="b35e4-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-134">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-136">
        - Selection</span></span><br><span data-ttu-id="b35e4-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-137">
        - Settings</span></span><br><span data-ttu-id="b35e4-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-138">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-139">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-140">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-142">Office on Windows</span></span><br><span data-ttu-id="b35e4-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-144">- TaskPane</span></span><br><span data-ttu-id="b35e4-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-145">
        - Content</span></span><br><span data-ttu-id="b35e4-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-146">
        - Custom Functions</span></span><br><span data-ttu-id="b35e4-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b35e4-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b35e4-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b35e4-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b35e4-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b35e4-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-161">
        - BindingEvents</span></span><br><span data-ttu-id="b35e4-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-162">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-163">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-164">
        - File</span></span><br><span data-ttu-id="b35e4-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-165">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-167">
        - Selection</span></span><br><span data-ttu-id="b35e4-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-168">
        - Settings</span></span><br><span data-ttu-id="b35e4-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-169">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-170">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-171">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-173">Office 2019 on Windows</span></span><br><span data-ttu-id="b35e4-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b35e4-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-175">- TaskPane</span></span><br><span data-ttu-id="b35e4-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-176">
        - Content</span></span><br><span data-ttu-id="b35e4-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b35e4-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-188">- BindingEvents</span></span><br><span data-ttu-id="b35e4-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-189">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-190">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-191">
        - File</span></span><br><span data-ttu-id="b35e4-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-192">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-194">
        - Selection</span></span><br><span data-ttu-id="b35e4-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-195">
        - Settings</span></span><br><span data-ttu-id="b35e4-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-196">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-197">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-198">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-200">Office 2016 on Windows</span></span><br><span data-ttu-id="b35e4-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b35e4-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-202">- TaskPane</span></span><br><span data-ttu-id="b35e4-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-203">
        - Content</span></span></td>
    <td><span data-ttu-id="b35e4-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b35e4-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-207">- BindingEvents</span></span><br><span data-ttu-id="b35e4-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-208">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-209">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-210">
        - File</span></span><br><span data-ttu-id="b35e4-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-211">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-213">
        - Selection</span></span><br><span data-ttu-id="b35e4-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-214">
        - Settings</span></span><br><span data-ttu-id="b35e4-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-215">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-216">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-217">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-219">Office 2013 on Windows</span></span><br><span data-ttu-id="b35e4-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b35e4-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-221">
        - TaskPane</span></span><br><span data-ttu-id="b35e4-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b35e4-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b35e4-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b35e4-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-225">
        - BindingEvents</span></span><br><span data-ttu-id="b35e4-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-226">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-227">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-228">
        - File</span></span><br><span data-ttu-id="b35e4-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-229">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-231">
        - Selection</span></span><br><span data-ttu-id="b35e4-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-232">
        - Settings</span></span><br><span data-ttu-id="b35e4-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-233">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-234">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-235">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b35e4-237">Office on iPad</span></span><br><span data-ttu-id="b35e4-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b35e4-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-239">- TaskPane</span></span><br><span data-ttu-id="b35e4-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-240">
        - Content</span></span></td>
    <td><span data-ttu-id="b35e4-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b35e4-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b35e4-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-253">- BindingEvents</span></span><br><span data-ttu-id="b35e4-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-254">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-255">
        - File</span></span><br><span data-ttu-id="b35e4-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-256">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-258">
        - Selection</span></span><br><span data-ttu-id="b35e4-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-259">
        - Settings</span></span><br><span data-ttu-id="b35e4-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-260">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-261">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-262">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-264">Office on Mac</span></span><br><span data-ttu-id="b35e4-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b35e4-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-266">- TaskPane</span></span><br><span data-ttu-id="b35e4-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-267">
        - Content</span></span><br><span data-ttu-id="b35e4-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-268">
        - Custom Functions</span></span><br><span data-ttu-id="b35e4-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b35e4-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b35e4-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b35e4-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b35e4-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-283">- BindingEvents</span></span><br><span data-ttu-id="b35e4-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-284">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-285">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-286">
        - File</span></span><br><span data-ttu-id="b35e4-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-287">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-289">
        - PdfFile</span></span><br><span data-ttu-id="b35e4-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-290">
        - Selection</span></span><br><span data-ttu-id="b35e4-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-291">
        - Settings</span></span><br><span data-ttu-id="b35e4-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-292">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-293">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-294">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-296">Office 2019 on Mac</span></span><br><span data-ttu-id="b35e4-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b35e4-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-298">- TaskPane</span></span><br><span data-ttu-id="b35e4-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-299">
        - Content</span></span><br><span data-ttu-id="b35e4-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b35e4-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b35e4-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b35e4-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b35e4-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b35e4-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b35e4-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b35e4-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b35e4-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-311">- BindingEvents</span></span><br><span data-ttu-id="b35e4-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-312">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-313">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-314">
        - File</span></span><br><span data-ttu-id="b35e4-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-315">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-317">
        - PdfFile</span></span><br><span data-ttu-id="b35e4-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-318">
        - Selection</span></span><br><span data-ttu-id="b35e4-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-319">
        - Settings</span></span><br><span data-ttu-id="b35e4-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-320">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-321">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-322">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-324">Office 2016 on Mac</span></span><br><span data-ttu-id="b35e4-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b35e4-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-326">- TaskPane</span></span><br><span data-ttu-id="b35e4-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-327">
        - Content</span></span></td>
    <td><span data-ttu-id="b35e4-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b35e4-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b35e4-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b35e4-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-331">- BindingEvents</span></span><br><span data-ttu-id="b35e4-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-332">
        - CompressedFile</span></span><br><span data-ttu-id="b35e4-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-333">
        - DocumentEvents</span></span><br><span data-ttu-id="b35e4-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-334">
        - File</span></span><br><span data-ttu-id="b35e4-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-335">
        - MatrixBindings</span></span><br><span data-ttu-id="b35e4-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-337">
        - PdfFile</span></span><br><span data-ttu-id="b35e4-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-338">
        - Selection</span></span><br><span data-ttu-id="b35e4-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-339">
        - Settings</span></span><br><span data-ttu-id="b35e4-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-340">
        - TableBindings</span></span><br><span data-ttu-id="b35e4-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-341">
        - TableCoercion</span></span><br><span data-ttu-id="b35e4-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-342">
        - TextBindings</span></span><br><span data-ttu-id="b35e4-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b35e4-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b35e4-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="b35e4-345">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="b35e4-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b35e4-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="b35e4-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b35e4-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b35e4-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b35e4-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-350">Office on the web</span></span></td>
    <td><span data-ttu-id="b35e4-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b35e4-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-353">Office on Windows</span></span><br><span data-ttu-id="b35e4-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b35e4-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b35e4-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-357">Office for Mac</span></span><br><span data-ttu-id="b35e4-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b35e4-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b35e4-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b35e4-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b35e4-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="b35e4-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b35e4-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b35e4-362">Platform</span></span></th>
    <th><span data-ttu-id="b35e4-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-363">Extension points</span></span></th>
    <th><span data-ttu-id="b35e4-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="b35e4-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-366">Office on the web</span></span><br><span data-ttu-id="b35e4-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="b35e4-367">(modern)</span></span></td>
    <td> <span data-ttu-id="b35e4-368">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-368">- Message Read</span></span><br><span data-ttu-id="b35e4-369">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-369">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-370">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-370">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-371">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-371">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b35e4-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b35e4-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b35e4-381">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-382">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-382">Office on the web</span></span><br><span data-ttu-id="b35e4-383">(classique)</span><span class="sxs-lookup"><span data-stu-id="b35e4-383">(classic)</span></span></td>
    <td> <span data-ttu-id="b35e4-384">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-384">- Message Read</span></span><br><span data-ttu-id="b35e4-385">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-385">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-386">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-386">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-387">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-387">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b35e4-395">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-396">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-396">Office on Windows</span></span><br><span data-ttu-id="b35e4-397">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-398">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-398">- Message Read</span></span><br><span data-ttu-id="b35e4-399">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-399">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-400">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-400">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-401">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-401">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b35e4-403">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b35e4-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b35e4-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b35e4-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b35e4-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b35e4-412">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-413">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-413">Office 2019 on Windows</span></span><br><span data-ttu-id="b35e4-414">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-415">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-415">- Message Read</span></span><br><span data-ttu-id="b35e4-416">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-416">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-417">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-417">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-418">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-418">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b35e4-420">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b35e4-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b35e4-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b35e4-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b35e4-428">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-429">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-429">Office 2016 on Windows</span></span><br><span data-ttu-id="b35e4-430">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-431">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-431">- Message Read</span></span><br><span data-ttu-id="b35e4-432">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-432">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-433">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-433">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-434">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-434">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b35e4-436">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b35e4-436">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b35e4-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b35e4-441">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-442">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-442">Office 2013 on Windows</span></span><br><span data-ttu-id="b35e4-443">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-444">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-444">- Message Read</span></span><br><span data-ttu-id="b35e4-445">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-445">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-446">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-446">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-447">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-447">
      - Appointment Organizer (Compose)</span></span><br>
    <td> <span data-ttu-id="b35e4-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b35e4-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b35e4-452">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-453">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="b35e4-453">Office on iOS</span></span><br><span data-ttu-id="b35e4-454">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-455">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-455">- Message Read</span></span><br><span data-ttu-id="b35e4-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b35e4-462">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-463">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-463">Office on Mac</span></span><br><span data-ttu-id="b35e4-464">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-465">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-465">- Message Read</span></span><br><span data-ttu-id="b35e4-466">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-466">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-467">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-467">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-468">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-468">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b35e4-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b35e4-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b35e4-478">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-479">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-479">Office 2019 on Mac</span></span><br><span data-ttu-id="b35e4-480">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-481">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-481">- Message Read</span></span><br><span data-ttu-id="b35e4-482">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-482">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-483">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-483">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-484">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-484">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b35e4-492">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-493">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-493">Office 2016 on Mac</span></span><br><span data-ttu-id="b35e4-494">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-495">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-495">- Message Read</span></span><br><span data-ttu-id="b35e4-496">
      - Composer un message</span><span class="sxs-lookup"><span data-stu-id="b35e4-496">
      - Message Compose</span></span><br><span data-ttu-id="b35e4-497">
      - Participant au rendez-vous (lecture)</span><span class="sxs-lookup"><span data-stu-id="b35e4-497">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b35e4-498">
      - Organisateur de rendez-vous (composer)</span><span class="sxs-lookup"><span data-stu-id="b35e4-498">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b35e4-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b35e4-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b35e4-506">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-507">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="b35e4-507">Office on Android</span></span><br><span data-ttu-id="b35e4-508">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-509">- Message lu</span><span class="sxs-lookup"><span data-stu-id="b35e4-509">- Message Read</span></span><br><span data-ttu-id="b35e4-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b35e4-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b35e4-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b35e4-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b35e4-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b35e4-516">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b35e4-516">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b35e4-517">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b35e4-517">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b35e4-518">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b35e4-518">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b35e4-519">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="b35e4-519">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b35e4-520">Word</span><span class="sxs-lookup"><span data-stu-id="b35e4-520">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b35e4-521">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b35e4-521">Platform</span></span></th>
    <th><span data-ttu-id="b35e4-522">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-522">Extension points</span></span></th>
    <th><span data-ttu-id="b35e4-523">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-523">API requirement sets</span></span></th>
    <th><span data-ttu-id="b35e4-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-525">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-525">Office on the web</span></span></td>
    <td> <span data-ttu-id="b35e4-526">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-526">- TaskPane</span></span><br><span data-ttu-id="b35e4-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b35e4-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-534">- BindingEvents</span></span><br><span data-ttu-id="b35e4-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-536">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-537">
         - File</span></span><br><span data-ttu-id="b35e4-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-539">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-542">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-543">
         - Selection</span></span><br><span data-ttu-id="b35e4-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-544">
         - Settings</span></span><br><span data-ttu-id="b35e4-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-545">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-546">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-547">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-548">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-549">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-550">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-550">Office on Windows</span></span><br><span data-ttu-id="b35e4-551">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-551">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-552">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-552">- TaskPane</span></span><br><span data-ttu-id="b35e4-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b35e4-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-560">- BindingEvents</span></span><br><span data-ttu-id="b35e4-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-561">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-563">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-564">
         - File</span></span><br><span data-ttu-id="b35e4-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-566">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-569">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-570">
         - Selection</span></span><br><span data-ttu-id="b35e4-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-571">
         - Settings</span></span><br><span data-ttu-id="b35e4-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-572">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-573">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-574">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-575">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-577">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-577">Office 2019 on Windows</span></span><br><span data-ttu-id="b35e4-578">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-579">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="b35e4-579">- TaskPane</span></span><br><span data-ttu-id="b35e4-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-586">- BindingEvents</span></span><br><span data-ttu-id="b35e4-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-587">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-589">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-590">
         - File</span></span><br><span data-ttu-id="b35e4-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-592">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-595">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-596">
         - Selection</span></span><br><span data-ttu-id="b35e4-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-597">
         - Settings</span></span><br><span data-ttu-id="b35e4-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-598">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-599">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-600">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-601">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-603">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-603">Office 2016 on Windows</span></span><br><span data-ttu-id="b35e4-604">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-605">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b35e4-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-609">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-609">- BindingEvents</span></span><br><span data-ttu-id="b35e4-610">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-610">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-611">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-611">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-612">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-613">
         - File</span></span><br><span data-ttu-id="b35e4-614">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-614">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-615">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-615">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-616">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-616">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-617">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-617">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-618">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-618">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-619">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-619">
         - Selection</span></span><br><span data-ttu-id="b35e4-620">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-620">
         - Settings</span></span><br><span data-ttu-id="b35e4-621">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-621">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-622">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-622">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-623">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-623">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-624">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-625">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-625">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-626">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-626">Office 2013 on Windows</span></span><br><span data-ttu-id="b35e4-627">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-627">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-628">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-628">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b35e4-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b35e4-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-631">- BindingEvents</span></span><br><span data-ttu-id="b35e4-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-632">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-634">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-635">
         - File</span></span><br><span data-ttu-id="b35e4-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-637">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-640">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-641">
         - Selection</span></span><br><span data-ttu-id="b35e4-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-642">
         - Settings</span></span><br><span data-ttu-id="b35e4-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-643">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-644">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-645">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-646">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-647">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-648">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b35e4-648">Office on iPad</span></span><br><span data-ttu-id="b35e4-649">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-650">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-650">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b35e4-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-656">- BindingEvents</span></span><br><span data-ttu-id="b35e4-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-657">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-659">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-660">
         - File</span></span><br><span data-ttu-id="b35e4-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-662">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-665">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-666">
         - Selection</span></span><br><span data-ttu-id="b35e4-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-667">
         - Settings</span></span><br><span data-ttu-id="b35e4-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-668">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-669">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-670">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-671">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-673">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-673">Office on Mac</span></span><br><span data-ttu-id="b35e4-674">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-674">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-675">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-675">- TaskPane</span></span><br><span data-ttu-id="b35e4-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b35e4-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-683">- BindingEvents</span></span><br><span data-ttu-id="b35e4-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-684">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-686">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-687">
         - File</span></span><br><span data-ttu-id="b35e4-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-689">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-692">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-693">
         - Selection</span></span><br><span data-ttu-id="b35e4-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-694">
         - Settings</span></span><br><span data-ttu-id="b35e4-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-695">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-696">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-697">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-698">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-700">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-700">Office 2019 on Mac</span></span><br><span data-ttu-id="b35e4-701">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-702">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="b35e4-702">- TaskPane</span></span><br><span data-ttu-id="b35e4-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b35e4-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b35e4-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b35e4-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-709">- BindingEvents</span></span><br><span data-ttu-id="b35e4-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-710">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-712">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-713">
         - File</span></span><br><span data-ttu-id="b35e4-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-715">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-718">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-719">
         - Selection</span></span><br><span data-ttu-id="b35e4-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-720">
         - Settings</span></span><br><span data-ttu-id="b35e4-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-721">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-722">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-723">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-724">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-725">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-726">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-726">Office 2016 on Mac</span></span><br><span data-ttu-id="b35e4-727">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-727">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-728">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-728">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b35e4-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b35e4-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b35e4-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-732">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-732">- BindingEvents</span></span><br><span data-ttu-id="b35e4-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-733">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-734">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b35e4-734">
         - CustomXmlParts</span></span><br><span data-ttu-id="b35e4-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-735">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-736">
         - File</span></span><br><span data-ttu-id="b35e4-737">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-737">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-738">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-738">
         - MatrixBindings</span></span><br><span data-ttu-id="b35e4-739">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-739">
         - MatrixCoercion</span></span><br><span data-ttu-id="b35e4-740">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-740">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b35e4-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-741">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-742">
         - Selection</span></span><br><span data-ttu-id="b35e4-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-743">
         - Settings</span></span><br><span data-ttu-id="b35e4-744">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-744">
         - TableBindings</span></span><br><span data-ttu-id="b35e4-745">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-745">
         - TableCoercion</span></span><br><span data-ttu-id="b35e4-746">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b35e4-746">
         - TextBindings</span></span><br><span data-ttu-id="b35e4-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-747">
         - TextCoercion</span></span><br><span data-ttu-id="b35e4-748">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-748">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b35e4-749">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b35e4-749">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b35e4-750">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b35e4-750">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b35e4-751">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b35e4-751">Platform</span></span></th>
    <th><span data-ttu-id="b35e4-752">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-752">Extension points</span></span></th>
    <th><span data-ttu-id="b35e4-753">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-753">API requirement sets</span></span></th>
    <th><span data-ttu-id="b35e4-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-755">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-755">Office on the web</span></span></td>
    <td> <span data-ttu-id="b35e4-756">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-756">- Content</span></span><br><span data-ttu-id="b35e4-757">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-757">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b35e4-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b35e4-763">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-763">- ActiveView</span></span><br><span data-ttu-id="b35e4-764">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-764">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-765">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-765">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-766">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-766">
         - File</span></span><br><span data-ttu-id="b35e4-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-767">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-768">
         - Selection</span></span><br><span data-ttu-id="b35e4-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-769">
         - Settings</span></span><br><span data-ttu-id="b35e4-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-771">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-771">Office on Windows</span></span><br><span data-ttu-id="b35e4-772">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-772">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-773">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-773">- Content</span></span><br><span data-ttu-id="b35e4-774">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-774">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b35e4-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b35e4-780">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-780">- ActiveView</span></span><br><span data-ttu-id="b35e4-781">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-781">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-782">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-782">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-783">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-783">
         - File</span></span><br><span data-ttu-id="b35e4-784">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-784">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-785">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-785">
         - Selection</span></span><br><span data-ttu-id="b35e4-786">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-786">
         - Settings</span></span><br><span data-ttu-id="b35e4-787">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-787">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-788">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-788">Office 2019 on Windows</span></span><br><span data-ttu-id="b35e4-789">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-789">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-790">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-790">- Content</span></span><br><span data-ttu-id="b35e4-791">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-791">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-795">- ActiveView</span></span><br><span data-ttu-id="b35e4-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-796">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-797">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-798">
         - File</span></span><br><span data-ttu-id="b35e4-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-799">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-800">
         - Selection</span></span><br><span data-ttu-id="b35e4-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-801">
         - Settings</span></span><br><span data-ttu-id="b35e4-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-803">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-803">Office 2016 on Windows</span></span><br><span data-ttu-id="b35e4-804">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-804">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-805">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-805">- Content</span></span><br><span data-ttu-id="b35e4-806">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b35e4-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b35e4-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-809">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-809">- ActiveView</span></span><br><span data-ttu-id="b35e4-810">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-810">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-811">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-811">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-812">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-812">
         - File</span></span><br><span data-ttu-id="b35e4-813">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-813">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-814">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-814">
         - Selection</span></span><br><span data-ttu-id="b35e4-815">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-815">
         - Settings</span></span><br><span data-ttu-id="b35e4-816">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-816">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-817">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-817">Office 2013 on Windows</span></span><br><span data-ttu-id="b35e4-818">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-818">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-819">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-819">- Content</span></span><br><span data-ttu-id="b35e4-820">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-820">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b35e4-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b35e4-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b35e4-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-823">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-823">- ActiveView</span></span><br><span data-ttu-id="b35e4-824">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-824">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-825">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-825">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-826">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-826">
         - File</span></span><br><span data-ttu-id="b35e4-827">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-827">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-828">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-828">
         - Selection</span></span><br><span data-ttu-id="b35e4-829">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-829">
         - Settings</span></span><br><span data-ttu-id="b35e4-830">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-830">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-831">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b35e4-831">Office on iPad</span></span><br><span data-ttu-id="b35e4-832">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-832">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-833">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-833">- Content</span></span><br><span data-ttu-id="b35e4-834">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-834">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b35e4-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-838">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-838">- ActiveView</span></span><br><span data-ttu-id="b35e4-839">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-839">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-840">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-840">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-841">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-841">
         - File</span></span><br><span data-ttu-id="b35e4-842">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-842">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-843">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-843">
         - Selection</span></span><br><span data-ttu-id="b35e4-844">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-844">
         - Settings</span></span><br><span data-ttu-id="b35e4-845">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-845">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-846">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-846">Office on Mac</span></span><br><span data-ttu-id="b35e4-847">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b35e4-847">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b35e4-848">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-848">- Content</span></span><br><span data-ttu-id="b35e4-849">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-849">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b35e4-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b35e4-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b35e4-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-855">- ActiveView</span></span><br><span data-ttu-id="b35e4-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-856">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-857">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-858">
         - File</span></span><br><span data-ttu-id="b35e4-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-859">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-860">
         - Selection</span></span><br><span data-ttu-id="b35e4-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-861">
         - Settings</span></span><br><span data-ttu-id="b35e4-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-862">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-863">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-863">Office 2019 on Mac</span></span><br><span data-ttu-id="b35e4-864">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-864">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-865">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-865">- Content</span></span><br><span data-ttu-id="b35e4-866">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-866">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-870">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-870">- ActiveView</span></span><br><span data-ttu-id="b35e4-871">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-871">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-872">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-872">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-873">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-873">
         - File</span></span><br><span data-ttu-id="b35e4-874">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-874">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-875">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-875">
         - Selection</span></span><br><span data-ttu-id="b35e4-876">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-876">
         - Settings</span></span><br><span data-ttu-id="b35e4-877">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-877">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-878">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-878">Office 2016 on Mac</span></span><br><span data-ttu-id="b35e4-879">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-879">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-880">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-880">- Content</span></span><br><span data-ttu-id="b35e4-881">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-881">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b35e4-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b35e4-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-884">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b35e4-884">- ActiveView</span></span><br><span data-ttu-id="b35e4-885">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-885">
         - CompressedFile</span></span><br><span data-ttu-id="b35e4-886">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-886">
         - DocumentEvents</span></span><br><span data-ttu-id="b35e4-887">
         - File</span><span class="sxs-lookup"><span data-stu-id="b35e4-887">
         - File</span></span><br><span data-ttu-id="b35e4-888">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b35e4-888">
         - PdfFile</span></span><br><span data-ttu-id="b35e4-889">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-889">
         - Selection</span></span><br><span data-ttu-id="b35e4-890">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-890">
         - Settings</span></span><br><span data-ttu-id="b35e4-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b35e4-892">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b35e4-892">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b35e4-893">OneNote</span><span class="sxs-lookup"><span data-stu-id="b35e4-893">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b35e4-894">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b35e4-894">Platform</span></span></th>
    <th><span data-ttu-id="b35e4-895">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-895">Extension points</span></span></th>
    <th><span data-ttu-id="b35e4-896">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-896">API requirement sets</span></span></th>
    <th><span data-ttu-id="b35e4-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-898">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b35e4-898">Office on the web</span></span></td>
    <td> <span data-ttu-id="b35e4-899">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b35e4-899">- Content</span></span><br><span data-ttu-id="b35e4-900">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-900">
         - TaskPane</span></span><br><span data-ttu-id="b35e4-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b35e4-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b35e4-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b35e4-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-905">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b35e4-905">- DocumentEvents</span></span><br><span data-ttu-id="b35e4-906">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-906">
         - HtmlCoercion</span></span><br><span data-ttu-id="b35e4-907">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b35e4-907">
         - Settings</span></span><br><span data-ttu-id="b35e4-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b35e4-909">Projet</span><span class="sxs-lookup"><span data-stu-id="b35e4-909">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b35e4-910">Plateforme</span><span class="sxs-lookup"><span data-stu-id="b35e4-910">Platform</span></span></th>
    <th><span data-ttu-id="b35e4-911">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b35e4-911">Extension points</span></span></th>
    <th><span data-ttu-id="b35e4-912">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-912">API requirement sets</span></span></th>
    <th><span data-ttu-id="b35e4-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b35e4-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-914">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-914">Office 2019 on Windows</span></span><br><span data-ttu-id="b35e4-915">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-915">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-916">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-916">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-918">- Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-918">- Selection</span></span><br><span data-ttu-id="b35e4-919">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-919">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-920">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-920">Office 2016 on Windows</span></span><br><span data-ttu-id="b35e4-921">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-921">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-922">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-922">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-924">- Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-924">- Selection</span></span><br><span data-ttu-id="b35e4-925">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-925">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b35e4-926">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b35e4-926">Office 2013 on Windows</span></span><br><span data-ttu-id="b35e4-927">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b35e4-927">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b35e4-928">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-928">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b35e4-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b35e4-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b35e4-930">- Selection</span><span class="sxs-lookup"><span data-stu-id="b35e4-930">- Selection</span></span><br><span data-ttu-id="b35e4-931">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b35e4-931">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b35e4-932">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b35e4-932">See also</span></span>

- [<span data-ttu-id="b35e4-933">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-933">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b35e4-934">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b35e4-934">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b35e4-935">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="b35e4-935">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b35e4-936">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="b35e4-936">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b35e4-937">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="b35e4-937">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b35e4-938">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="b35e4-938">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b35e4-939">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="b35e4-939">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b35e4-940">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="b35e4-940">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b35e4-941">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b35e4-941">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b35e4-942">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b35e4-942">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b35e4-943">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="b35e4-943">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="b35e4-944">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="b35e4-944">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)