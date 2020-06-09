---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611981"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="39045-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="39045-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="39045-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="39045-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="39045-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="39045-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="39045-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="39045-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="39045-108">Excel</span><span class="sxs-lookup"><span data-stu-id="39045-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="39045-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="39045-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="39045-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="39045-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="39045-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="39045-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-114">- TaskPane</span></span><br><span data-ttu-id="39045-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-115">
        - Content</span></span><br><span data-ttu-id="39045-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-116">
        - Custom Functions</span></span><br><span data-ttu-id="39045-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="39045-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="39045-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39045-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39045-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="39045-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="39045-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="39045-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="39045-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="39045-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="39045-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-131">
        - BindingEvents</span></span><br><span data-ttu-id="39045-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-132">
        - CompressedFile</span></span><br><span data-ttu-id="39045-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-133">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-134">
        - File</span></span><br><span data-ttu-id="39045-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-135">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-137">
        - Selection</span></span><br><span data-ttu-id="39045-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-138">
        - Settings</span></span><br><span data-ttu-id="39045-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-139">
        - TableBindings</span></span><br><span data-ttu-id="39045-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-140">
        - TableCoercion</span></span><br><span data-ttu-id="39045-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-141">
        - TextBindings</span></span><br><span data-ttu-id="39045-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-143">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="39045-143">Office on Windows</span></span><br><span data-ttu-id="39045-144">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-145">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-145">- TaskPane</span></span><br><span data-ttu-id="39045-146">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-146">
        - Content</span></span><br><span data-ttu-id="39045-147">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-147">
        - Custom Functions</span></span><br><span data-ttu-id="39045-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="39045-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="39045-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39045-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39045-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="39045-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="39045-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="39045-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="39045-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="39045-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-163">
        - BindingEvents</span></span><br><span data-ttu-id="39045-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-164">
        - CompressedFile</span></span><br><span data-ttu-id="39045-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-165">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-166">
        - File</span></span><br><span data-ttu-id="39045-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-167">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-169">
        - Selection</span></span><br><span data-ttu-id="39045-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-170">
        - Settings</span></span><br><span data-ttu-id="39045-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-171">
        - TableBindings</span></span><br><span data-ttu-id="39045-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-172">
        - TableCoercion</span></span><br><span data-ttu-id="39045-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-173">
        - TextBindings</span></span><br><span data-ttu-id="39045-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-175">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-175">Office 2019 on Windows</span></span><br><span data-ttu-id="39045-176">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39045-177">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-177">- TaskPane</span></span><br><span data-ttu-id="39045-178">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-178">
        - Content</span></span><br><span data-ttu-id="39045-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39045-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-190">- BindingEvents</span></span><br><span data-ttu-id="39045-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-191">
        - CompressedFile</span></span><br><span data-ttu-id="39045-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-192">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-193">
        - File</span></span><br><span data-ttu-id="39045-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-194">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-196">
        - Selection</span></span><br><span data-ttu-id="39045-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-197">
        - Settings</span></span><br><span data-ttu-id="39045-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-198">
        - TableBindings</span></span><br><span data-ttu-id="39045-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-199">
        - TableCoercion</span></span><br><span data-ttu-id="39045-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-200">
        - TextBindings</span></span><br><span data-ttu-id="39045-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-202">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-202">Office 2016 on Windows</span></span><br><span data-ttu-id="39045-203">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39045-204">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-204">- TaskPane</span></span><br><span data-ttu-id="39045-205">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-205">
        - Content</span></span></td>
    <td><span data-ttu-id="39045-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39045-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-209">- BindingEvents</span></span><br><span data-ttu-id="39045-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-210">
        - CompressedFile</span></span><br><span data-ttu-id="39045-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-211">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-212">
        - File</span></span><br><span data-ttu-id="39045-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-213">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-215">
        - Selection</span></span><br><span data-ttu-id="39045-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-216">
        - Settings</span></span><br><span data-ttu-id="39045-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-217">
        - TableBindings</span></span><br><span data-ttu-id="39045-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-218">
        - TableCoercion</span></span><br><span data-ttu-id="39045-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-219">
        - TextBindings</span></span><br><span data-ttu-id="39045-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-221">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-221">Office 2013 on Windows</span></span><br><span data-ttu-id="39045-222">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39045-223">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-223">
        - TaskPane</span></span><br><span data-ttu-id="39045-224">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="39045-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39045-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39045-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-227">
        - BindingEvents</span></span><br><span data-ttu-id="39045-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-228">
        - CompressedFile</span></span><br><span data-ttu-id="39045-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-229">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-230">
        - File</span></span><br><span data-ttu-id="39045-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-231">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-233">
        - Selection</span></span><br><span data-ttu-id="39045-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-234">
        - Settings</span></span><br><span data-ttu-id="39045-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-235">
        - TableBindings</span></span><br><span data-ttu-id="39045-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-236">
        - TableCoercion</span></span><br><span data-ttu-id="39045-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-237">
        - TextBindings</span></span><br><span data-ttu-id="39045-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-239">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="39045-239">Office on iPad</span></span><br><span data-ttu-id="39045-240">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39045-241">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-241">- TaskPane</span></span><br><span data-ttu-id="39045-242">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-242">
        - Content</span></span></td>
    <td><span data-ttu-id="39045-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39045-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39045-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="39045-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="39045-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="39045-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="39045-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-256">- BindingEvents</span></span><br><span data-ttu-id="39045-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-257">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-258">
        - File</span></span><br><span data-ttu-id="39045-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-259">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-261">
        - Selection</span></span><br><span data-ttu-id="39045-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-262">
        - Settings</span></span><br><span data-ttu-id="39045-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-263">
        - TableBindings</span></span><br><span data-ttu-id="39045-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-264">
        - TableCoercion</span></span><br><span data-ttu-id="39045-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-265">
        - TextBindings</span></span><br><span data-ttu-id="39045-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-267">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-267">Office on Mac</span></span><br><span data-ttu-id="39045-268">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39045-269">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-269">- TaskPane</span></span><br><span data-ttu-id="39045-270">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-270">
        - Content</span></span><br><span data-ttu-id="39045-271">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-271">
        - Custom Functions</span></span><br><span data-ttu-id="39045-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39045-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39045-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39045-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="39045-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="39045-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="39045-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="39045-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="39045-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-287">- BindingEvents</span></span><br><span data-ttu-id="39045-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-288">
        - CompressedFile</span></span><br><span data-ttu-id="39045-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-289">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-290">
        - File</span></span><br><span data-ttu-id="39045-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-291">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-293">
        - PdfFile</span></span><br><span data-ttu-id="39045-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-294">
        - Selection</span></span><br><span data-ttu-id="39045-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-295">
        - Settings</span></span><br><span data-ttu-id="39045-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-296">
        - TableBindings</span></span><br><span data-ttu-id="39045-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-297">
        - TableCoercion</span></span><br><span data-ttu-id="39045-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-298">
        - TextBindings</span></span><br><span data-ttu-id="39045-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-300">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-300">Office 2019 on Mac</span></span><br><span data-ttu-id="39045-301">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39045-302">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-302">- TaskPane</span></span><br><span data-ttu-id="39045-303">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-303">
        - Content</span></span><br><span data-ttu-id="39045-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39045-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39045-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39045-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39045-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39045-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39045-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39045-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39045-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-315">- BindingEvents</span></span><br><span data-ttu-id="39045-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-316">
        - CompressedFile</span></span><br><span data-ttu-id="39045-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-317">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-318">
        - File</span></span><br><span data-ttu-id="39045-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-319">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-321">
        - PdfFile</span></span><br><span data-ttu-id="39045-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-322">
        - Selection</span></span><br><span data-ttu-id="39045-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-323">
        - Settings</span></span><br><span data-ttu-id="39045-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-324">
        - TableBindings</span></span><br><span data-ttu-id="39045-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-325">
        - TableCoercion</span></span><br><span data-ttu-id="39045-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-326">
        - TextBindings</span></span><br><span data-ttu-id="39045-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-328">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-328">Office 2016 on Mac</span></span><br><span data-ttu-id="39045-329">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39045-330">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-330">- TaskPane</span></span><br><span data-ttu-id="39045-331">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-331">
        - Content</span></span></td>
    <td><span data-ttu-id="39045-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39045-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39045-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39045-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-335">- BindingEvents</span></span><br><span data-ttu-id="39045-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-336">
        - CompressedFile</span></span><br><span data-ttu-id="39045-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-337">
        - DocumentEvents</span></span><br><span data-ttu-id="39045-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="39045-338">
        - File</span></span><br><span data-ttu-id="39045-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-339">
        - MatrixBindings</span></span><br><span data-ttu-id="39045-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="39045-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-341">
        - PdfFile</span></span><br><span data-ttu-id="39045-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-342">
        - Selection</span></span><br><span data-ttu-id="39045-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-343">
        - Settings</span></span><br><span data-ttu-id="39045-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-344">
        - TableBindings</span></span><br><span data-ttu-id="39045-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-345">
        - TableCoercion</span></span><br><span data-ttu-id="39045-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-346">
        - TextBindings</span></span><br><span data-ttu-id="39045-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="39045-348">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="39045-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="39045-349">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="39045-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="39045-350">Plateforme</span><span class="sxs-lookup"><span data-stu-id="39045-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="39045-351">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="39045-352">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="39045-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-354">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-354">Office on the web</span></span></td>
    <td><span data-ttu-id="39045-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39045-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-357">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="39045-357">Office on Windows</span></span><br><span data-ttu-id="39045-358">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39045-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39045-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-361">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="39045-361">Office for Mac</span></span><br><span data-ttu-id="39045-362">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="39045-363">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="39045-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39045-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="39045-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="39045-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39045-366">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="39045-366">Platform</span></span></th>
    <th><span data-ttu-id="39045-367">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-367">Extension points</span></span></th>
    <th><span data-ttu-id="39045-368">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="39045-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-370">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-370">Office on the web</span></span><br><span data-ttu-id="39045-371">(moderne)</span><span class="sxs-lookup"><span data-stu-id="39045-371">(modern)</span></span></td>
    <td> <span data-ttu-id="39045-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39045-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39045-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39045-385">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-386">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-386">Office on the web</span></span><br><span data-ttu-id="39045-387">(classique)</span><span class="sxs-lookup"><span data-stu-id="39045-387">(classic)</span></span></td>
    <td> <span data-ttu-id="39045-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39045-399">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-400">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="39045-400">Office on Windows</span></span><br><span data-ttu-id="39045-401">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39045-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="39045-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="39045-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39045-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39045-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39045-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-417">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-417">Office 2019 on Windows</span></span><br><span data-ttu-id="39045-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39045-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="39045-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="39045-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39045-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="39045-432">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-433">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-433">Office 2016 on Windows</span></span><br><span data-ttu-id="39045-434">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39045-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="39045-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="39045-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="39045-445">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-446">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-446">Office 2013 on Windows</span></span><br><span data-ttu-id="39045-447">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="39045-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="39045-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="39045-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-457">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="39045-457">Office on iOS</span></span><br><span data-ttu-id="39045-458">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="39045-466">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-467">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-467">Office on Mac</span></span><br><span data-ttu-id="39045-468">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39045-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39045-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39045-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39045-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39045-482">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-483">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-483">Office 2019 on Mac</span></span><br><span data-ttu-id="39045-484">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39045-496">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-497">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-497">Office 2016 on Mac</span></span><br><span data-ttu-id="39045-498">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="39045-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="39045-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="39045-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="39045-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="39045-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="39045-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39045-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39045-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39045-510">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-511">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="39045-511">Office on Android</span></span><br><span data-ttu-id="39045-512">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="39045-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="39045-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)</span><span class="sxs-lookup"><span data-stu-id="39045-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="39045-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39045-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39045-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39045-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39045-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39045-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39045-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="39045-521">Non disponible</span><span class="sxs-lookup"><span data-stu-id="39045-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="39045-522">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="39045-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="39045-523">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="39045-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="39045-524">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="39045-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="39045-525">Word</span><span class="sxs-lookup"><span data-stu-id="39045-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39045-526">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="39045-526">Platform</span></span></th>
    <th><span data-ttu-id="39045-527">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-527">Extension points</span></span></th>
    <th><span data-ttu-id="39045-528">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="39045-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-530">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="39045-531">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-531">- TaskPane</span></span><br><span data-ttu-id="39045-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39045-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-539">- BindingEvents</span></span><br><span data-ttu-id="39045-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-541">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-542">
         - File</span></span><br><span data-ttu-id="39045-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-544">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-547">
         - PdfFile</span></span><br><span data-ttu-id="39045-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-548">
         - Selection</span></span><br><span data-ttu-id="39045-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-549">
         - Settings</span></span><br><span data-ttu-id="39045-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-550">
         - TableBindings</span></span><br><span data-ttu-id="39045-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-551">
         - TableCoercion</span></span><br><span data-ttu-id="39045-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-552">
         - TextBindings</span></span><br><span data-ttu-id="39045-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-553">
         - TextCoercion</span></span><br><span data-ttu-id="39045-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-555">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="39045-555">Office on Windows</span></span><br><span data-ttu-id="39045-556">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-557">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-557">- TaskPane</span></span><br><span data-ttu-id="39045-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39045-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-565">- BindingEvents</span></span><br><span data-ttu-id="39045-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-566">
         - CompressedFile</span></span><br><span data-ttu-id="39045-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-568">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-569">
         - File</span></span><br><span data-ttu-id="39045-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-571">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-574">
         - PdfFile</span></span><br><span data-ttu-id="39045-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-575">
         - Selection</span></span><br><span data-ttu-id="39045-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-576">
         - Settings</span></span><br><span data-ttu-id="39045-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-577">
         - TableBindings</span></span><br><span data-ttu-id="39045-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-578">
         - TableCoercion</span></span><br><span data-ttu-id="39045-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-579">
         - TextBindings</span></span><br><span data-ttu-id="39045-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-580">
         - TextCoercion</span></span><br><span data-ttu-id="39045-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-582">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-582">Office 2019 on Windows</span></span><br><span data-ttu-id="39045-583">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-584">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="39045-584">- TaskPane</span></span><br><span data-ttu-id="39045-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-591">- BindingEvents</span></span><br><span data-ttu-id="39045-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-592">
         - CompressedFile</span></span><br><span data-ttu-id="39045-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-594">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-595">
         - File</span></span><br><span data-ttu-id="39045-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-597">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-600">
         - PdfFile</span></span><br><span data-ttu-id="39045-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-601">
         - Selection</span></span><br><span data-ttu-id="39045-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-602">
         - Settings</span></span><br><span data-ttu-id="39045-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-603">
         - TableBindings</span></span><br><span data-ttu-id="39045-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-604">
         - TableCoercion</span></span><br><span data-ttu-id="39045-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-605">
         - TextBindings</span></span><br><span data-ttu-id="39045-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-606">
         - TextCoercion</span></span><br><span data-ttu-id="39045-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-608">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-608">Office 2016 on Windows</span></span><br><span data-ttu-id="39045-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39045-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-614">- BindingEvents</span></span><br><span data-ttu-id="39045-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-615">
         - CompressedFile</span></span><br><span data-ttu-id="39045-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-617">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-618">
         - File</span></span><br><span data-ttu-id="39045-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-620">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-623">
         - PdfFile</span></span><br><span data-ttu-id="39045-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-624">
         - Selection</span></span><br><span data-ttu-id="39045-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-625">
         - Settings</span></span><br><span data-ttu-id="39045-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-626">
         - TableBindings</span></span><br><span data-ttu-id="39045-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-627">
         - TableCoercion</span></span><br><span data-ttu-id="39045-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-628">
         - TextBindings</span></span><br><span data-ttu-id="39045-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-629">
         - TextCoercion</span></span><br><span data-ttu-id="39045-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-631">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-631">Office 2013 on Windows</span></span><br><span data-ttu-id="39045-632">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-633">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39045-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39045-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-636">- BindingEvents</span></span><br><span data-ttu-id="39045-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-637">
         - CompressedFile</span></span><br><span data-ttu-id="39045-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-639">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-640">
         - File</span></span><br><span data-ttu-id="39045-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-642">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-645">
         - PdfFile</span></span><br><span data-ttu-id="39045-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-646">
         - Selection</span></span><br><span data-ttu-id="39045-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-647">
         - Settings</span></span><br><span data-ttu-id="39045-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-648">
         - TableBindings</span></span><br><span data-ttu-id="39045-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-649">
         - TableCoercion</span></span><br><span data-ttu-id="39045-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-650">
         - TextBindings</span></span><br><span data-ttu-id="39045-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-651">
         - TextCoercion</span></span><br><span data-ttu-id="39045-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-653">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="39045-653">Office on iPad</span></span><br><span data-ttu-id="39045-654">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-655">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="39045-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-661">- BindingEvents</span></span><br><span data-ttu-id="39045-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-662">
         - CompressedFile</span></span><br><span data-ttu-id="39045-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-664">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-665">
         - File</span></span><br><span data-ttu-id="39045-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-667">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-670">
         - PdfFile</span></span><br><span data-ttu-id="39045-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-671">
         - Selection</span></span><br><span data-ttu-id="39045-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-672">
         - Settings</span></span><br><span data-ttu-id="39045-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-673">
         - TableBindings</span></span><br><span data-ttu-id="39045-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-674">
         - TableCoercion</span></span><br><span data-ttu-id="39045-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-675">
         - TextBindings</span></span><br><span data-ttu-id="39045-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-676">
         - TextCoercion</span></span><br><span data-ttu-id="39045-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-678">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-678">Office on Mac</span></span><br><span data-ttu-id="39045-679">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-680">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-680">- TaskPane</span></span><br><span data-ttu-id="39045-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="39045-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-688">- BindingEvents</span></span><br><span data-ttu-id="39045-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-689">
         - CompressedFile</span></span><br><span data-ttu-id="39045-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-691">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-692">
         - File</span></span><br><span data-ttu-id="39045-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-694">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-697">
         - PdfFile</span></span><br><span data-ttu-id="39045-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-698">
         - Selection</span></span><br><span data-ttu-id="39045-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-699">
         - Settings</span></span><br><span data-ttu-id="39045-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-700">
         - TableBindings</span></span><br><span data-ttu-id="39045-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-701">
         - TableCoercion</span></span><br><span data-ttu-id="39045-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-702">
         - TextBindings</span></span><br><span data-ttu-id="39045-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-703">
         - TextCoercion</span></span><br><span data-ttu-id="39045-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-705">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-705">Office 2019 on Mac</span></span><br><span data-ttu-id="39045-706">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-707">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="39045-707">- TaskPane</span></span><br><span data-ttu-id="39045-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39045-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39045-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39045-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="39045-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-714">- BindingEvents</span></span><br><span data-ttu-id="39045-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-715">
         - CompressedFile</span></span><br><span data-ttu-id="39045-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-717">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-718">
         - File</span></span><br><span data-ttu-id="39045-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-720">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-723">
         - PdfFile</span></span><br><span data-ttu-id="39045-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-724">
         - Selection</span></span><br><span data-ttu-id="39045-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-725">
         - Settings</span></span><br><span data-ttu-id="39045-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-726">
         - TableBindings</span></span><br><span data-ttu-id="39045-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-727">
         - TableCoercion</span></span><br><span data-ttu-id="39045-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-728">
         - TextBindings</span></span><br><span data-ttu-id="39045-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-729">
         - TextCoercion</span></span><br><span data-ttu-id="39045-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-731">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-731">Office 2016 on Mac</span></span><br><span data-ttu-id="39045-732">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-733">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39045-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39045-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39045-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39045-737">- BindingEvents</span></span><br><span data-ttu-id="39045-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-738">
         - CompressedFile</span></span><br><span data-ttu-id="39045-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39045-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="39045-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-740">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-741">
         - File</span></span><br><span data-ttu-id="39045-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39045-743">
         - MatrixBindings</span></span><br><span data-ttu-id="39045-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="39045-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39045-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-746">
         - PdfFile</span></span><br><span data-ttu-id="39045-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-747">
         - Selection</span></span><br><span data-ttu-id="39045-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-748">
         - Settings</span></span><br><span data-ttu-id="39045-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39045-749">
         - TableBindings</span></span><br><span data-ttu-id="39045-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-750">
         - TableCoercion</span></span><br><span data-ttu-id="39045-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39045-751">
         - TextBindings</span></span><br><span data-ttu-id="39045-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-752">
         - TextCoercion</span></span><br><span data-ttu-id="39045-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39045-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="39045-754">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="39045-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="39045-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="39045-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39045-756">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="39045-756">Platform</span></span></th>
    <th><span data-ttu-id="39045-757">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-757">Extension points</span></span></th>
    <th><span data-ttu-id="39045-758">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="39045-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-760">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="39045-761">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-761">- Content</span></span><br><span data-ttu-id="39045-762">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-762">
         - TaskPane</span></span><br><span data-ttu-id="39045-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39045-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39045-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-768">- ActiveView</span></span><br><span data-ttu-id="39045-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-769">
         - CompressedFile</span></span><br><span data-ttu-id="39045-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-770">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-771">
         - File</span></span><br><span data-ttu-id="39045-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-772">
         - PdfFile</span></span><br><span data-ttu-id="39045-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-773">
         - Selection</span></span><br><span data-ttu-id="39045-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-774">
         - Settings</span></span><br><span data-ttu-id="39045-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-776">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="39045-776">Office on Windows</span></span><br><span data-ttu-id="39045-777">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-778">- Content</span></span><br><span data-ttu-id="39045-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-779">
         - TaskPane</span></span><br><span data-ttu-id="39045-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39045-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39045-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-785">- ActiveView</span></span><br><span data-ttu-id="39045-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-786">
         - CompressedFile</span></span><br><span data-ttu-id="39045-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-787">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-788">
         - File</span></span><br><span data-ttu-id="39045-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-789">
         - PdfFile</span></span><br><span data-ttu-id="39045-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-790">
         - Selection</span></span><br><span data-ttu-id="39045-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-791">
         - Settings</span></span><br><span data-ttu-id="39045-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-793">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-793">Office 2019 on Windows</span></span><br><span data-ttu-id="39045-794">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-795">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-795">- Content</span></span><br><span data-ttu-id="39045-796">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-796">
         - TaskPane</span></span><br><span data-ttu-id="39045-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-800">- ActiveView</span></span><br><span data-ttu-id="39045-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-801">
         - CompressedFile</span></span><br><span data-ttu-id="39045-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-802">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-803">
         - File</span></span><br><span data-ttu-id="39045-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-804">
         - PdfFile</span></span><br><span data-ttu-id="39045-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-805">
         - Selection</span></span><br><span data-ttu-id="39045-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-806">
         - Settings</span></span><br><span data-ttu-id="39045-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-808">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-808">Office 2016 on Windows</span></span><br><span data-ttu-id="39045-809">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-810">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-810">- Content</span></span><br><span data-ttu-id="39045-811">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39045-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39045-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-814">- ActiveView</span></span><br><span data-ttu-id="39045-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-815">
         - CompressedFile</span></span><br><span data-ttu-id="39045-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-816">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-817">
         - File</span></span><br><span data-ttu-id="39045-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-818">
         - PdfFile</span></span><br><span data-ttu-id="39045-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-819">
         - Selection</span></span><br><span data-ttu-id="39045-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-820">
         - Settings</span></span><br><span data-ttu-id="39045-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-822">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-822">Office 2013 on Windows</span></span><br><span data-ttu-id="39045-823">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-824">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-824">- Content</span></span><br><span data-ttu-id="39045-825">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="39045-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39045-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39045-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-828">- ActiveView</span></span><br><span data-ttu-id="39045-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-829">
         - CompressedFile</span></span><br><span data-ttu-id="39045-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-830">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-831">
         - File</span></span><br><span data-ttu-id="39045-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-832">
         - PdfFile</span></span><br><span data-ttu-id="39045-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-833">
         - Selection</span></span><br><span data-ttu-id="39045-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-834">
         - Settings</span></span><br><span data-ttu-id="39045-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-836">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="39045-836">Office on iPad</span></span><br><span data-ttu-id="39045-837">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-838">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-838">- Content</span></span><br><span data-ttu-id="39045-839">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39045-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-843">- ActiveView</span></span><br><span data-ttu-id="39045-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-844">
         - CompressedFile</span></span><br><span data-ttu-id="39045-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-845">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-846">
         - File</span></span><br><span data-ttu-id="39045-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-847">
         - PdfFile</span></span><br><span data-ttu-id="39045-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-848">
         - Selection</span></span><br><span data-ttu-id="39045-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-849">
         - Settings</span></span><br><span data-ttu-id="39045-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-851">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-851">Office on Mac</span></span><br><span data-ttu-id="39045-852">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="39045-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39045-853">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-853">- Content</span></span><br><span data-ttu-id="39045-854">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-854">
         - TaskPane</span></span><br><span data-ttu-id="39045-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39045-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39045-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39045-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39045-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-860">- ActiveView</span></span><br><span data-ttu-id="39045-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-861">
         - CompressedFile</span></span><br><span data-ttu-id="39045-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-862">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-863">
         - File</span></span><br><span data-ttu-id="39045-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-864">
         - PdfFile</span></span><br><span data-ttu-id="39045-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-865">
         - Selection</span></span><br><span data-ttu-id="39045-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-866">
         - Settings</span></span><br><span data-ttu-id="39045-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-868">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-868">Office 2019 on Mac</span></span><br><span data-ttu-id="39045-869">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-870">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-870">- Content</span></span><br><span data-ttu-id="39045-871">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-871">
         - TaskPane</span></span><br><span data-ttu-id="39045-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-875">- ActiveView</span></span><br><span data-ttu-id="39045-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-876">
         - CompressedFile</span></span><br><span data-ttu-id="39045-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-877">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-878">
         - File</span></span><br><span data-ttu-id="39045-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-879">
         - PdfFile</span></span><br><span data-ttu-id="39045-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-880">
         - Selection</span></span><br><span data-ttu-id="39045-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-881">
         - Settings</span></span><br><span data-ttu-id="39045-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-883">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="39045-883">Office 2016 on Mac</span></span><br><span data-ttu-id="39045-884">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-885">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-885">- Content</span></span><br><span data-ttu-id="39045-886">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39045-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39045-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39045-889">- ActiveView</span></span><br><span data-ttu-id="39045-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39045-890">
         - CompressedFile</span></span><br><span data-ttu-id="39045-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-891">
         - DocumentEvents</span></span><br><span data-ttu-id="39045-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="39045-892">
         - File</span></span><br><span data-ttu-id="39045-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39045-893">
         - PdfFile</span></span><br><span data-ttu-id="39045-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39045-894">
         - Selection</span></span><br><span data-ttu-id="39045-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-895">
         - Settings</span></span><br><span data-ttu-id="39045-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="39045-897">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="39045-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="39045-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="39045-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39045-899">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="39045-899">Platform</span></span></th>
    <th><span data-ttu-id="39045-900">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-900">Extension points</span></span></th>
    <th><span data-ttu-id="39045-901">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="39045-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-903">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="39045-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="39045-904">- Contenu</span><span class="sxs-lookup"><span data-stu-id="39045-904">- Content</span></span><br><span data-ttu-id="39045-905">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-905">
         - TaskPane</span></span><br><span data-ttu-id="39045-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="39045-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39045-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="39045-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39045-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39045-910">- DocumentEvents</span></span><br><span data-ttu-id="39045-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="39045-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39045-912">
         - Settings</span></span><br><span data-ttu-id="39045-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="39045-914">Projet</span><span class="sxs-lookup"><span data-stu-id="39045-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39045-915">Plateforme</span><span class="sxs-lookup"><span data-stu-id="39045-915">Platform</span></span></th>
    <th><span data-ttu-id="39045-916">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="39045-916">Extension points</span></span></th>
    <th><span data-ttu-id="39045-917">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="39045-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="39045-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-919">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-919">Office 2019 on Windows</span></span><br><span data-ttu-id="39045-920">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-921">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="39045-923">- Selection</span></span><br><span data-ttu-id="39045-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-925">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-925">Office 2016 on Windows</span></span><br><span data-ttu-id="39045-926">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-927">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="39045-929">- Selection</span></span><br><span data-ttu-id="39045-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39045-931">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="39045-931">Office 2013 on Windows</span></span><br><span data-ttu-id="39045-932">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="39045-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39045-933">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="39045-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39045-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39045-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39045-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="39045-935">- Selection</span></span><br><span data-ttu-id="39045-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39045-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="39045-937">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="39045-937">See also</span></span>

- [<span data-ttu-id="39045-938">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="39045-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="39045-939">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="39045-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="39045-940">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="39045-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="39045-941">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="39045-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="39045-942">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="39045-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="39045-943">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="39045-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="39045-944">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="39045-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="39045-945">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="39045-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="39045-946">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="39045-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="39045-947">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="39045-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="39045-948">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="39045-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="39045-949">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="39045-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)