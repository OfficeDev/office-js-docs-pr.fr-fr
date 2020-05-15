---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217822"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="564f7-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="564f7-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="564f7-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="564f7-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="564f7-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="564f7-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="564f7-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="564f7-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="564f7-108">Excel</span><span class="sxs-lookup"><span data-stu-id="564f7-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="564f7-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="564f7-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="564f7-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="564f7-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="564f7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="564f7-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-114">- TaskPane</span></span><br><span data-ttu-id="564f7-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-115">
        - Content</span></span><br><span data-ttu-id="564f7-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-116">
        - Custom Functions</span></span><br><span data-ttu-id="564f7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="564f7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="564f7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="564f7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="564f7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="564f7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="564f7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="564f7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="564f7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="564f7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="564f7-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-131">
        - BindingEvents</span></span><br><span data-ttu-id="564f7-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-132">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-133">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-134">
        - File</span></span><br><span data-ttu-id="564f7-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-135">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-137">
        - Selection</span></span><br><span data-ttu-id="564f7-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-138">
        - Settings</span></span><br><span data-ttu-id="564f7-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-139">
        - TableBindings</span></span><br><span data-ttu-id="564f7-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-140">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-141">
        - TextBindings</span></span><br><span data-ttu-id="564f7-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-143">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-143">Office on Windows</span></span><br><span data-ttu-id="564f7-144">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-145">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-145">- TaskPane</span></span><br><span data-ttu-id="564f7-146">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-146">
        - Content</span></span><br><span data-ttu-id="564f7-147">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-147">
        - Custom Functions</span></span><br><span data-ttu-id="564f7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="564f7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="564f7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="564f7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="564f7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="564f7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="564f7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="564f7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="564f7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="564f7-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-163">
        - BindingEvents</span></span><br><span data-ttu-id="564f7-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-164">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-165">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-166">
        - File</span></span><br><span data-ttu-id="564f7-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-167">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-169">
        - Selection</span></span><br><span data-ttu-id="564f7-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-170">
        - Settings</span></span><br><span data-ttu-id="564f7-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-171">
        - TableBindings</span></span><br><span data-ttu-id="564f7-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-172">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-173">
        - TextBindings</span></span><br><span data-ttu-id="564f7-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-175">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-175">Office 2019 on Windows</span></span><br><span data-ttu-id="564f7-176">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="564f7-177">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-177">- TaskPane</span></span><br><span data-ttu-id="564f7-178">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-178">
        - Content</span></span><br><span data-ttu-id="564f7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="564f7-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-190">- BindingEvents</span></span><br><span data-ttu-id="564f7-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-191">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-192">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-193">
        - File</span></span><br><span data-ttu-id="564f7-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-194">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-196">
        - Selection</span></span><br><span data-ttu-id="564f7-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-197">
        - Settings</span></span><br><span data-ttu-id="564f7-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-198">
        - TableBindings</span></span><br><span data-ttu-id="564f7-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-199">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-200">
        - TextBindings</span></span><br><span data-ttu-id="564f7-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-202">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-202">Office 2016 on Windows</span></span><br><span data-ttu-id="564f7-203">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="564f7-204">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-204">- TaskPane</span></span><br><span data-ttu-id="564f7-205">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-205">
        - Content</span></span></td>
    <td><span data-ttu-id="564f7-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="564f7-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-209">- BindingEvents</span></span><br><span data-ttu-id="564f7-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-210">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-211">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-212">
        - File</span></span><br><span data-ttu-id="564f7-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-213">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-215">
        - Selection</span></span><br><span data-ttu-id="564f7-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-216">
        - Settings</span></span><br><span data-ttu-id="564f7-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-217">
        - TableBindings</span></span><br><span data-ttu-id="564f7-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-218">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-219">
        - TextBindings</span></span><br><span data-ttu-id="564f7-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-221">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-221">Office 2013 on Windows</span></span><br><span data-ttu-id="564f7-222">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="564f7-223">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-223">
        - TaskPane</span></span><br><span data-ttu-id="564f7-224">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="564f7-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="564f7-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="564f7-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-227">
        - BindingEvents</span></span><br><span data-ttu-id="564f7-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-228">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-229">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-230">
        - File</span></span><br><span data-ttu-id="564f7-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-231">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-233">
        - Selection</span></span><br><span data-ttu-id="564f7-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-234">
        - Settings</span></span><br><span data-ttu-id="564f7-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-235">
        - TableBindings</span></span><br><span data-ttu-id="564f7-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-236">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-237">
        - TextBindings</span></span><br><span data-ttu-id="564f7-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-239">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="564f7-239">Office on iPad</span></span><br><span data-ttu-id="564f7-240">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="564f7-241">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-241">- TaskPane</span></span><br><span data-ttu-id="564f7-242">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-242">
        - Content</span></span></td>
    <td><span data-ttu-id="564f7-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="564f7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="564f7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="564f7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="564f7-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="564f7-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="564f7-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-256">- BindingEvents</span></span><br><span data-ttu-id="564f7-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-257">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-258">
        - File</span></span><br><span data-ttu-id="564f7-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-259">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-261">
        - Selection</span></span><br><span data-ttu-id="564f7-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-262">
        - Settings</span></span><br><span data-ttu-id="564f7-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-263">
        - TableBindings</span></span><br><span data-ttu-id="564f7-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-264">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-265">
        - TextBindings</span></span><br><span data-ttu-id="564f7-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-267">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-267">Office on Mac</span></span><br><span data-ttu-id="564f7-268">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="564f7-269">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-269">- TaskPane</span></span><br><span data-ttu-id="564f7-270">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-270">
        - Content</span></span><br><span data-ttu-id="564f7-271">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-271">
        - Custom Functions</span></span><br><span data-ttu-id="564f7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="564f7-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="564f7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="564f7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="564f7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="564f7-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="564f7-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="564f7-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="564f7-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-287">- BindingEvents</span></span><br><span data-ttu-id="564f7-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-288">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-289">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-290">
        - File</span></span><br><span data-ttu-id="564f7-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-291">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-293">
        - PdfFile</span></span><br><span data-ttu-id="564f7-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-294">
        - Selection</span></span><br><span data-ttu-id="564f7-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-295">
        - Settings</span></span><br><span data-ttu-id="564f7-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-296">
        - TableBindings</span></span><br><span data-ttu-id="564f7-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-297">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-298">
        - TextBindings</span></span><br><span data-ttu-id="564f7-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-300">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-300">Office 2019 on Mac</span></span><br><span data-ttu-id="564f7-301">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="564f7-302">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-302">- TaskPane</span></span><br><span data-ttu-id="564f7-303">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-303">
        - Content</span></span><br><span data-ttu-id="564f7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="564f7-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="564f7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="564f7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="564f7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="564f7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="564f7-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="564f7-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="564f7-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-315">- BindingEvents</span></span><br><span data-ttu-id="564f7-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-316">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-317">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-318">
        - File</span></span><br><span data-ttu-id="564f7-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-319">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-321">
        - PdfFile</span></span><br><span data-ttu-id="564f7-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-322">
        - Selection</span></span><br><span data-ttu-id="564f7-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-323">
        - Settings</span></span><br><span data-ttu-id="564f7-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-324">
        - TableBindings</span></span><br><span data-ttu-id="564f7-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-325">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-326">
        - TextBindings</span></span><br><span data-ttu-id="564f7-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-328">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-328">Office 2016 on Mac</span></span><br><span data-ttu-id="564f7-329">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="564f7-330">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-330">- TaskPane</span></span><br><span data-ttu-id="564f7-331">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-331">
        - Content</span></span></td>
    <td><span data-ttu-id="564f7-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="564f7-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="564f7-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="564f7-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-335">- BindingEvents</span></span><br><span data-ttu-id="564f7-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-336">
        - CompressedFile</span></span><br><span data-ttu-id="564f7-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-337">
        - DocumentEvents</span></span><br><span data-ttu-id="564f7-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="564f7-338">
        - File</span></span><br><span data-ttu-id="564f7-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-339">
        - MatrixBindings</span></span><br><span data-ttu-id="564f7-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="564f7-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-341">
        - PdfFile</span></span><br><span data-ttu-id="564f7-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-342">
        - Selection</span></span><br><span data-ttu-id="564f7-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-343">
        - Settings</span></span><br><span data-ttu-id="564f7-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-344">
        - TableBindings</span></span><br><span data-ttu-id="564f7-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-345">
        - TableCoercion</span></span><br><span data-ttu-id="564f7-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-346">
        - TextBindings</span></span><br><span data-ttu-id="564f7-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="564f7-348">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="564f7-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="564f7-349">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="564f7-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="564f7-350">Plateforme</span><span class="sxs-lookup"><span data-stu-id="564f7-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="564f7-351">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="564f7-352">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="564f7-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-354">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-354">Office on the web</span></span></td>
    <td><span data-ttu-id="564f7-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="564f7-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-357">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-357">Office on Windows</span></span><br><span data-ttu-id="564f7-358">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="564f7-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="564f7-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-361">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-361">Office for Mac</span></span><br><span data-ttu-id="564f7-362">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="564f7-363">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="564f7-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="564f7-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="564f7-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="564f7-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="564f7-366">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="564f7-366">Platform</span></span></th>
    <th><span data-ttu-id="564f7-367">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-367">Extension points</span></span></th>
    <th><span data-ttu-id="564f7-368">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="564f7-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-370">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-370">Office on the web</span></span><br><span data-ttu-id="564f7-371">(moderne)</span><span class="sxs-lookup"><span data-stu-id="564f7-371">(modern)</span></span></td>
    <td> <span data-ttu-id="564f7-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="564f7-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="564f7-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="564f7-385">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-386">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-386">Office on the web</span></span><br><span data-ttu-id="564f7-387">(classique)</span><span class="sxs-lookup"><span data-stu-id="564f7-387">(classic)</span></span></td>
    <td> <span data-ttu-id="564f7-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="564f7-399">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-400">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-400">Office on Windows</span></span><br><span data-ttu-id="564f7-401">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="564f7-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="564f7-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="564f7-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="564f7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="564f7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="564f7-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-417">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-417">Office 2019 on Windows</span></span><br><span data-ttu-id="564f7-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="564f7-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="564f7-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="564f7-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="564f7-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="564f7-432">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-433">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-433">Office 2016 on Windows</span></span><br><span data-ttu-id="564f7-434">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="564f7-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="564f7-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="564f7-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="564f7-445">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-446">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-446">Office 2013 on Windows</span></span><br><span data-ttu-id="564f7-447">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="564f7-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="564f7-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="564f7-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-457">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="564f7-457">Office on iOS</span></span><br><span data-ttu-id="564f7-458">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="564f7-466">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-467">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-467">Office on Mac</span></span><br><span data-ttu-id="564f7-468">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="564f7-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="564f7-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="564f7-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="564f7-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="564f7-482">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-483">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-483">Office 2019 on Mac</span></span><br><span data-ttu-id="564f7-484">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="564f7-496">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-497">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-497">Office 2016 on Mac</span></span><br><span data-ttu-id="564f7-498">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="564f7-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="564f7-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="564f7-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="564f7-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="564f7-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="564f7-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="564f7-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="564f7-510">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-511">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="564f7-511">Office on Android</span></span><br><span data-ttu-id="564f7-512">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="564f7-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="564f7-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)</span><span class="sxs-lookup"><span data-stu-id="564f7-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="564f7-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="564f7-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="564f7-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="564f7-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="564f7-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="564f7-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="564f7-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="564f7-521">Non disponible</span><span class="sxs-lookup"><span data-stu-id="564f7-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="564f7-522">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="564f7-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="564f7-523">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="564f7-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="564f7-524">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="564f7-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="564f7-525">Word</span><span class="sxs-lookup"><span data-stu-id="564f7-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="564f7-526">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="564f7-526">Platform</span></span></th>
    <th><span data-ttu-id="564f7-527">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-527">Extension points</span></span></th>
    <th><span data-ttu-id="564f7-528">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="564f7-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-530">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="564f7-531">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-531">- TaskPane</span></span><br><span data-ttu-id="564f7-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="564f7-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-539">- BindingEvents</span></span><br><span data-ttu-id="564f7-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-541">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-542">
         - File</span></span><br><span data-ttu-id="564f7-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-544">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-547">
         - PdfFile</span></span><br><span data-ttu-id="564f7-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-548">
         - Selection</span></span><br><span data-ttu-id="564f7-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-549">
         - Settings</span></span><br><span data-ttu-id="564f7-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-550">
         - TableBindings</span></span><br><span data-ttu-id="564f7-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-551">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-552">
         - TextBindings</span></span><br><span data-ttu-id="564f7-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-553">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-555">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-555">Office on Windows</span></span><br><span data-ttu-id="564f7-556">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-557">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-557">- TaskPane</span></span><br><span data-ttu-id="564f7-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="564f7-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-565">- BindingEvents</span></span><br><span data-ttu-id="564f7-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-566">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-568">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-569">
         - File</span></span><br><span data-ttu-id="564f7-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-571">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-574">
         - PdfFile</span></span><br><span data-ttu-id="564f7-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-575">
         - Selection</span></span><br><span data-ttu-id="564f7-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-576">
         - Settings</span></span><br><span data-ttu-id="564f7-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-577">
         - TableBindings</span></span><br><span data-ttu-id="564f7-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-578">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-579">
         - TextBindings</span></span><br><span data-ttu-id="564f7-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-580">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-582">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-582">Office 2019 on Windows</span></span><br><span data-ttu-id="564f7-583">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-584">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="564f7-584">- TaskPane</span></span><br><span data-ttu-id="564f7-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-591">- BindingEvents</span></span><br><span data-ttu-id="564f7-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-592">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-594">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-595">
         - File</span></span><br><span data-ttu-id="564f7-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-597">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-600">
         - PdfFile</span></span><br><span data-ttu-id="564f7-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-601">
         - Selection</span></span><br><span data-ttu-id="564f7-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-602">
         - Settings</span></span><br><span data-ttu-id="564f7-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-603">
         - TableBindings</span></span><br><span data-ttu-id="564f7-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-604">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-605">
         - TextBindings</span></span><br><span data-ttu-id="564f7-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-606">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-608">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-608">Office 2016 on Windows</span></span><br><span data-ttu-id="564f7-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="564f7-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-614">- BindingEvents</span></span><br><span data-ttu-id="564f7-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-615">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-617">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-618">
         - File</span></span><br><span data-ttu-id="564f7-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-620">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-623">
         - PdfFile</span></span><br><span data-ttu-id="564f7-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-624">
         - Selection</span></span><br><span data-ttu-id="564f7-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-625">
         - Settings</span></span><br><span data-ttu-id="564f7-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-626">
         - TableBindings</span></span><br><span data-ttu-id="564f7-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-627">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-628">
         - TextBindings</span></span><br><span data-ttu-id="564f7-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-629">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-631">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-631">Office 2013 on Windows</span></span><br><span data-ttu-id="564f7-632">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-633">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="564f7-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="564f7-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-636">- BindingEvents</span></span><br><span data-ttu-id="564f7-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-637">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-639">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-640">
         - File</span></span><br><span data-ttu-id="564f7-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-642">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-645">
         - PdfFile</span></span><br><span data-ttu-id="564f7-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-646">
         - Selection</span></span><br><span data-ttu-id="564f7-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-647">
         - Settings</span></span><br><span data-ttu-id="564f7-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-648">
         - TableBindings</span></span><br><span data-ttu-id="564f7-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-649">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-650">
         - TextBindings</span></span><br><span data-ttu-id="564f7-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-651">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-653">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="564f7-653">Office on iPad</span></span><br><span data-ttu-id="564f7-654">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-655">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="564f7-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-661">- BindingEvents</span></span><br><span data-ttu-id="564f7-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-662">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-664">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-665">
         - File</span></span><br><span data-ttu-id="564f7-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-667">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-670">
         - PdfFile</span></span><br><span data-ttu-id="564f7-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-671">
         - Selection</span></span><br><span data-ttu-id="564f7-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-672">
         - Settings</span></span><br><span data-ttu-id="564f7-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-673">
         - TableBindings</span></span><br><span data-ttu-id="564f7-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-674">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-675">
         - TextBindings</span></span><br><span data-ttu-id="564f7-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-676">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-678">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-678">Office on Mac</span></span><br><span data-ttu-id="564f7-679">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-680">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-680">- TaskPane</span></span><br><span data-ttu-id="564f7-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="564f7-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-688">- BindingEvents</span></span><br><span data-ttu-id="564f7-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-689">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-691">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-692">
         - File</span></span><br><span data-ttu-id="564f7-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-694">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-697">
         - PdfFile</span></span><br><span data-ttu-id="564f7-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-698">
         - Selection</span></span><br><span data-ttu-id="564f7-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-699">
         - Settings</span></span><br><span data-ttu-id="564f7-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-700">
         - TableBindings</span></span><br><span data-ttu-id="564f7-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-701">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-702">
         - TextBindings</span></span><br><span data-ttu-id="564f7-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-703">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-705">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-705">Office 2019 on Mac</span></span><br><span data-ttu-id="564f7-706">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-707">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="564f7-707">- TaskPane</span></span><br><span data-ttu-id="564f7-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="564f7-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="564f7-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="564f7-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="564f7-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-714">- BindingEvents</span></span><br><span data-ttu-id="564f7-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-715">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-717">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-718">
         - File</span></span><br><span data-ttu-id="564f7-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-720">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-723">
         - PdfFile</span></span><br><span data-ttu-id="564f7-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-724">
         - Selection</span></span><br><span data-ttu-id="564f7-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-725">
         - Settings</span></span><br><span data-ttu-id="564f7-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-726">
         - TableBindings</span></span><br><span data-ttu-id="564f7-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-727">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-728">
         - TextBindings</span></span><br><span data-ttu-id="564f7-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-729">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-731">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-731">Office 2016 on Mac</span></span><br><span data-ttu-id="564f7-732">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-733">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="564f7-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="564f7-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="564f7-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-737">- BindingEvents</span></span><br><span data-ttu-id="564f7-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-738">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="564f7-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="564f7-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-740">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-741">
         - File</span></span><br><span data-ttu-id="564f7-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-743">
         - MatrixBindings</span></span><br><span data-ttu-id="564f7-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="564f7-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="564f7-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-746">
         - PdfFile</span></span><br><span data-ttu-id="564f7-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-747">
         - Selection</span></span><br><span data-ttu-id="564f7-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-748">
         - Settings</span></span><br><span data-ttu-id="564f7-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-749">
         - TableBindings</span></span><br><span data-ttu-id="564f7-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-750">
         - TableCoercion</span></span><br><span data-ttu-id="564f7-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="564f7-751">
         - TextBindings</span></span><br><span data-ttu-id="564f7-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-752">
         - TextCoercion</span></span><br><span data-ttu-id="564f7-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="564f7-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="564f7-754">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="564f7-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="564f7-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="564f7-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="564f7-756">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="564f7-756">Platform</span></span></th>
    <th><span data-ttu-id="564f7-757">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-757">Extension points</span></span></th>
    <th><span data-ttu-id="564f7-758">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="564f7-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-760">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="564f7-761">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-761">- Content</span></span><br><span data-ttu-id="564f7-762">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-762">
         - TaskPane</span></span><br><span data-ttu-id="564f7-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="564f7-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="564f7-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-768">- ActiveView</span></span><br><span data-ttu-id="564f7-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-769">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-770">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-771">
         - File</span></span><br><span data-ttu-id="564f7-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-772">
         - PdfFile</span></span><br><span data-ttu-id="564f7-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-773">
         - Selection</span></span><br><span data-ttu-id="564f7-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-774">
         - Settings</span></span><br><span data-ttu-id="564f7-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-776">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-776">Office on Windows</span></span><br><span data-ttu-id="564f7-777">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-778">- Content</span></span><br><span data-ttu-id="564f7-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-779">
         - TaskPane</span></span><br><span data-ttu-id="564f7-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="564f7-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="564f7-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-785">- ActiveView</span></span><br><span data-ttu-id="564f7-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-786">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-787">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-788">
         - File</span></span><br><span data-ttu-id="564f7-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-789">
         - PdfFile</span></span><br><span data-ttu-id="564f7-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-790">
         - Selection</span></span><br><span data-ttu-id="564f7-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-791">
         - Settings</span></span><br><span data-ttu-id="564f7-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-793">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-793">Office 2019 on Windows</span></span><br><span data-ttu-id="564f7-794">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-795">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-795">- Content</span></span><br><span data-ttu-id="564f7-796">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-796">
         - TaskPane</span></span><br><span data-ttu-id="564f7-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-800">- ActiveView</span></span><br><span data-ttu-id="564f7-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-801">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-802">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-803">
         - File</span></span><br><span data-ttu-id="564f7-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-804">
         - PdfFile</span></span><br><span data-ttu-id="564f7-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-805">
         - Selection</span></span><br><span data-ttu-id="564f7-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-806">
         - Settings</span></span><br><span data-ttu-id="564f7-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-808">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-808">Office 2016 on Windows</span></span><br><span data-ttu-id="564f7-809">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-810">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-810">- Content</span></span><br><span data-ttu-id="564f7-811">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="564f7-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="564f7-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-814">- ActiveView</span></span><br><span data-ttu-id="564f7-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-815">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-816">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-817">
         - File</span></span><br><span data-ttu-id="564f7-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-818">
         - PdfFile</span></span><br><span data-ttu-id="564f7-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-819">
         - Selection</span></span><br><span data-ttu-id="564f7-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-820">
         - Settings</span></span><br><span data-ttu-id="564f7-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-822">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-822">Office 2013 on Windows</span></span><br><span data-ttu-id="564f7-823">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-824">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-824">- Content</span></span><br><span data-ttu-id="564f7-825">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="564f7-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="564f7-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="564f7-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-828">- ActiveView</span></span><br><span data-ttu-id="564f7-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-829">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-830">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-831">
         - File</span></span><br><span data-ttu-id="564f7-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-832">
         - PdfFile</span></span><br><span data-ttu-id="564f7-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-833">
         - Selection</span></span><br><span data-ttu-id="564f7-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-834">
         - Settings</span></span><br><span data-ttu-id="564f7-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-836">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="564f7-836">Office on iPad</span></span><br><span data-ttu-id="564f7-837">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-838">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-838">- Content</span></span><br><span data-ttu-id="564f7-839">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="564f7-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-843">- ActiveView</span></span><br><span data-ttu-id="564f7-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-844">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-845">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-846">
         - File</span></span><br><span data-ttu-id="564f7-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-847">
         - PdfFile</span></span><br><span data-ttu-id="564f7-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-848">
         - Selection</span></span><br><span data-ttu-id="564f7-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-849">
         - Settings</span></span><br><span data-ttu-id="564f7-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-851">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-851">Office on Mac</span></span><br><span data-ttu-id="564f7-852">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="564f7-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="564f7-853">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-853">- Content</span></span><br><span data-ttu-id="564f7-854">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-854">
         - TaskPane</span></span><br><span data-ttu-id="564f7-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="564f7-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="564f7-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="564f7-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="564f7-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-860">- ActiveView</span></span><br><span data-ttu-id="564f7-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-861">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-862">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-863">
         - File</span></span><br><span data-ttu-id="564f7-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-864">
         - PdfFile</span></span><br><span data-ttu-id="564f7-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-865">
         - Selection</span></span><br><span data-ttu-id="564f7-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-866">
         - Settings</span></span><br><span data-ttu-id="564f7-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-868">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-868">Office 2019 on Mac</span></span><br><span data-ttu-id="564f7-869">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-870">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-870">- Content</span></span><br><span data-ttu-id="564f7-871">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-871">
         - TaskPane</span></span><br><span data-ttu-id="564f7-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-875">- ActiveView</span></span><br><span data-ttu-id="564f7-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-876">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-877">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-878">
         - File</span></span><br><span data-ttu-id="564f7-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-879">
         - PdfFile</span></span><br><span data-ttu-id="564f7-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-880">
         - Selection</span></span><br><span data-ttu-id="564f7-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-881">
         - Settings</span></span><br><span data-ttu-id="564f7-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-883">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-883">Office 2016 on Mac</span></span><br><span data-ttu-id="564f7-884">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-885">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-885">- Content</span></span><br><span data-ttu-id="564f7-886">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="564f7-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="564f7-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="564f7-889">- ActiveView</span></span><br><span data-ttu-id="564f7-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="564f7-890">
         - CompressedFile</span></span><br><span data-ttu-id="564f7-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-891">
         - DocumentEvents</span></span><br><span data-ttu-id="564f7-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="564f7-892">
         - File</span></span><br><span data-ttu-id="564f7-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="564f7-893">
         - PdfFile</span></span><br><span data-ttu-id="564f7-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-894">
         - Selection</span></span><br><span data-ttu-id="564f7-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-895">
         - Settings</span></span><br><span data-ttu-id="564f7-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="564f7-897">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="564f7-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="564f7-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="564f7-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="564f7-899">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="564f7-899">Platform</span></span></th>
    <th><span data-ttu-id="564f7-900">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-900">Extension points</span></span></th>
    <th><span data-ttu-id="564f7-901">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="564f7-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-903">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="564f7-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="564f7-904">- Contenu</span><span class="sxs-lookup"><span data-stu-id="564f7-904">- Content</span></span><br><span data-ttu-id="564f7-905">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-905">
         - TaskPane</span></span><br><span data-ttu-id="564f7-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="564f7-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="564f7-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="564f7-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="564f7-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="564f7-910">- DocumentEvents</span></span><br><span data-ttu-id="564f7-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="564f7-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="564f7-912">
         - Settings</span></span><br><span data-ttu-id="564f7-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="564f7-914">Projet</span><span class="sxs-lookup"><span data-stu-id="564f7-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="564f7-915">Plateforme</span><span class="sxs-lookup"><span data-stu-id="564f7-915">Platform</span></span></th>
    <th><span data-ttu-id="564f7-916">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="564f7-916">Extension points</span></span></th>
    <th><span data-ttu-id="564f7-917">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="564f7-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="564f7-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-919">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-919">Office 2019 on Windows</span></span><br><span data-ttu-id="564f7-920">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-921">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-923">- Selection</span></span><br><span data-ttu-id="564f7-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-925">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-925">Office 2016 on Windows</span></span><br><span data-ttu-id="564f7-926">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-927">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-929">- Selection</span></span><br><span data-ttu-id="564f7-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="564f7-931">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="564f7-931">Office 2013 on Windows</span></span><br><span data-ttu-id="564f7-932">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="564f7-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="564f7-933">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="564f7-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="564f7-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="564f7-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="564f7-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="564f7-935">- Selection</span></span><br><span data-ttu-id="564f7-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="564f7-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="564f7-937">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="564f7-937">See also</span></span>

- [<span data-ttu-id="564f7-938">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="564f7-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="564f7-939">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="564f7-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="564f7-940">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="564f7-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="564f7-941">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="564f7-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="564f7-942">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="564f7-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="564f7-943">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="564f7-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="564f7-944">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="564f7-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="564f7-945">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="564f7-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="564f7-946">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="564f7-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="564f7-947">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="564f7-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="564f7-948">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="564f7-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="564f7-949">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="564f7-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)