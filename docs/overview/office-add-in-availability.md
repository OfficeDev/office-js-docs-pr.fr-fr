---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 04/13/2020
localization_priority: Priority
ms.openlocfilehash: 72da8db755fe6d1d166f66a70c8c298e5a27adff
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241055"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="df454-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="df454-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="df454-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="df454-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="df454-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="df454-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="df454-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="df454-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="df454-108">Excel</span><span class="sxs-lookup"><span data-stu-id="df454-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="df454-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="df454-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="df454-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="df454-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="df454-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="df454-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-114">- TaskPane</span></span><br><span data-ttu-id="df454-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-115">
        - Content</span></span><br><span data-ttu-id="df454-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-116">
        - Custom Functions</span></span><br><span data-ttu-id="df454-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="df454-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="df454-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="df454-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="df454-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="df454-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="df454-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="df454-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="df454-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-130">
        - BindingEvents</span></span><br><span data-ttu-id="df454-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-131">
        - CompressedFile</span></span><br><span data-ttu-id="df454-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-132">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-133">
        - File</span></span><br><span data-ttu-id="df454-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-134">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-136">
        - Selection</span></span><br><span data-ttu-id="df454-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-137">
        - Settings</span></span><br><span data-ttu-id="df454-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-138">
        - TableBindings</span></span><br><span data-ttu-id="df454-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-139">
        - TableCoercion</span></span><br><span data-ttu-id="df454-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-140">
        - TextBindings</span></span><br><span data-ttu-id="df454-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="df454-142">Office on Windows</span></span><br><span data-ttu-id="df454-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-144">- TaskPane</span></span><br><span data-ttu-id="df454-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-145">
        - Content</span></span><br><span data-ttu-id="df454-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-146">
        - Custom Functions</span></span><br><span data-ttu-id="df454-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="df454-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="df454-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="df454-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="df454-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="df454-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="df454-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="df454-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-161">
        - BindingEvents</span></span><br><span data-ttu-id="df454-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-162">
        - CompressedFile</span></span><br><span data-ttu-id="df454-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-163">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-164">
        - File</span></span><br><span data-ttu-id="df454-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-165">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-167">
        - Selection</span></span><br><span data-ttu-id="df454-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-168">
        - Settings</span></span><br><span data-ttu-id="df454-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-169">
        - TableBindings</span></span><br><span data-ttu-id="df454-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-170">
        - TableCoercion</span></span><br><span data-ttu-id="df454-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-171">
        - TextBindings</span></span><br><span data-ttu-id="df454-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-173">Office 2019 on Windows</span></span><br><span data-ttu-id="df454-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="df454-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-175">- TaskPane</span></span><br><span data-ttu-id="df454-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-176">
        - Content</span></span><br><span data-ttu-id="df454-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="df454-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-188">- BindingEvents</span></span><br><span data-ttu-id="df454-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-189">
        - CompressedFile</span></span><br><span data-ttu-id="df454-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-190">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-191">
        - File</span></span><br><span data-ttu-id="df454-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-192">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-194">
        - Selection</span></span><br><span data-ttu-id="df454-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-195">
        - Settings</span></span><br><span data-ttu-id="df454-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-196">
        - TableBindings</span></span><br><span data-ttu-id="df454-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-197">
        - TableCoercion</span></span><br><span data-ttu-id="df454-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-198">
        - TextBindings</span></span><br><span data-ttu-id="df454-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-200">Office 2016 on Windows</span></span><br><span data-ttu-id="df454-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="df454-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-202">- TaskPane</span></span><br><span data-ttu-id="df454-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-203">
        - Content</span></span></td>
    <td><span data-ttu-id="df454-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="df454-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-207">- BindingEvents</span></span><br><span data-ttu-id="df454-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-208">
        - CompressedFile</span></span><br><span data-ttu-id="df454-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-209">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-210">
        - File</span></span><br><span data-ttu-id="df454-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-211">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-213">
        - Selection</span></span><br><span data-ttu-id="df454-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-214">
        - Settings</span></span><br><span data-ttu-id="df454-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-215">
        - TableBindings</span></span><br><span data-ttu-id="df454-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-216">
        - TableCoercion</span></span><br><span data-ttu-id="df454-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-217">
        - TextBindings</span></span><br><span data-ttu-id="df454-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-219">Office 2013 on Windows</span></span><br><span data-ttu-id="df454-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="df454-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-221">
        - TaskPane</span></span><br><span data-ttu-id="df454-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="df454-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="df454-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="df454-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-225">
        - BindingEvents</span></span><br><span data-ttu-id="df454-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-226">
        - CompressedFile</span></span><br><span data-ttu-id="df454-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-227">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-228">
        - File</span></span><br><span data-ttu-id="df454-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-229">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-231">
        - Selection</span></span><br><span data-ttu-id="df454-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-232">
        - Settings</span></span><br><span data-ttu-id="df454-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-233">
        - TableBindings</span></span><br><span data-ttu-id="df454-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-234">
        - TableCoercion</span></span><br><span data-ttu-id="df454-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-235">
        - TextBindings</span></span><br><span data-ttu-id="df454-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="df454-237">Office on iPad</span></span><br><span data-ttu-id="df454-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="df454-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-239">- TaskPane</span></span><br><span data-ttu-id="df454-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-240">
        - Content</span></span></td>
    <td><span data-ttu-id="df454-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="df454-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="df454-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="df454-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="df454-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-253">- BindingEvents</span></span><br><span data-ttu-id="df454-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-254">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-255">
        - File</span></span><br><span data-ttu-id="df454-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-256">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-258">
        - Selection</span></span><br><span data-ttu-id="df454-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-259">
        - Settings</span></span><br><span data-ttu-id="df454-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-260">
        - TableBindings</span></span><br><span data-ttu-id="df454-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-261">
        - TableCoercion</span></span><br><span data-ttu-id="df454-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-262">
        - TextBindings</span></span><br><span data-ttu-id="df454-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-264">Office on Mac</span></span><br><span data-ttu-id="df454-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="df454-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-266">- TaskPane</span></span><br><span data-ttu-id="df454-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-267">
        - Content</span></span><br><span data-ttu-id="df454-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-268">
        - Custom Functions</span></span><br><span data-ttu-id="df454-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="df454-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="df454-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="df454-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="df454-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="df454-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="df454-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-283">- BindingEvents</span></span><br><span data-ttu-id="df454-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-284">
        - CompressedFile</span></span><br><span data-ttu-id="df454-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-285">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-286">
        - File</span></span><br><span data-ttu-id="df454-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-287">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-289">
        - PdfFile</span></span><br><span data-ttu-id="df454-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-290">
        - Selection</span></span><br><span data-ttu-id="df454-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-291">
        - Settings</span></span><br><span data-ttu-id="df454-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-292">
        - TableBindings</span></span><br><span data-ttu-id="df454-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-293">
        - TableCoercion</span></span><br><span data-ttu-id="df454-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-294">
        - TextBindings</span></span><br><span data-ttu-id="df454-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-296">Office 2019 on Mac</span></span><br><span data-ttu-id="df454-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="df454-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-298">- TaskPane</span></span><br><span data-ttu-id="df454-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-299">
        - Content</span></span><br><span data-ttu-id="df454-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="df454-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="df454-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="df454-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="df454-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="df454-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="df454-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="df454-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="df454-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-311">- BindingEvents</span></span><br><span data-ttu-id="df454-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-312">
        - CompressedFile</span></span><br><span data-ttu-id="df454-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-313">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-314">
        - File</span></span><br><span data-ttu-id="df454-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-315">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-317">
        - PdfFile</span></span><br><span data-ttu-id="df454-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-318">
        - Selection</span></span><br><span data-ttu-id="df454-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-319">
        - Settings</span></span><br><span data-ttu-id="df454-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-320">
        - TableBindings</span></span><br><span data-ttu-id="df454-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-321">
        - TableCoercion</span></span><br><span data-ttu-id="df454-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-322">
        - TextBindings</span></span><br><span data-ttu-id="df454-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-324">Office 2016 on Mac</span></span><br><span data-ttu-id="df454-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="df454-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-326">- TaskPane</span></span><br><span data-ttu-id="df454-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-327">
        - Content</span></span></td>
    <td><span data-ttu-id="df454-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="df454-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="df454-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="df454-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-331">- BindingEvents</span></span><br><span data-ttu-id="df454-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-332">
        - CompressedFile</span></span><br><span data-ttu-id="df454-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-333">
        - DocumentEvents</span></span><br><span data-ttu-id="df454-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="df454-334">
        - File</span></span><br><span data-ttu-id="df454-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-335">
        - MatrixBindings</span></span><br><span data-ttu-id="df454-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="df454-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-337">
        - PdfFile</span></span><br><span data-ttu-id="df454-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-338">
        - Selection</span></span><br><span data-ttu-id="df454-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-339">
        - Settings</span></span><br><span data-ttu-id="df454-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-340">
        - TableBindings</span></span><br><span data-ttu-id="df454-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-341">
        - TableCoercion</span></span><br><span data-ttu-id="df454-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-342">
        - TextBindings</span></span><br><span data-ttu-id="df454-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="df454-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="df454-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="df454-345">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="df454-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="df454-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="df454-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="df454-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="df454-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="df454-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-350">Office on the web</span></span></td>
    <td><span data-ttu-id="df454-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="df454-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="df454-353">Office on Windows</span></span><br><span data-ttu-id="df454-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="df454-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="df454-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="df454-357">Office for Mac</span></span><br><span data-ttu-id="df454-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="df454-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="df454-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="df454-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="df454-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="df454-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="df454-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="df454-362">Platform</span></span></th>
    <th><span data-ttu-id="df454-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-363">Extension points</span></span></th>
    <th><span data-ttu-id="df454-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="df454-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-366">Office on the web</span></span><br><span data-ttu-id="df454-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="df454-367">(modern)</span></span></td>
    <td> <span data-ttu-id="df454-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="df454-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="df454-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="df454-381">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-382">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-382">Office on the web</span></span><br><span data-ttu-id="df454-383">(classique)</span><span class="sxs-lookup"><span data-stu-id="df454-383">(classic)</span></span></td>
    <td> <span data-ttu-id="df454-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="df454-395">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-396">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="df454-396">Office on Windows</span></span><br><span data-ttu-id="df454-397">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="df454-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="df454-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="df454-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="df454-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="df454-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="df454-412">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-413">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-413">Office 2019 on Windows</span></span><br><span data-ttu-id="df454-414">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="df454-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="df454-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="df454-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="df454-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="df454-428">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-429">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-429">Office 2016 on Windows</span></span><br><span data-ttu-id="df454-430">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="df454-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="df454-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="df454-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="df454-441">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-442">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-442">Office 2013 on Windows</span></span><br><span data-ttu-id="df454-443">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="df454-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="df454-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="df454-452">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-453">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="df454-453">Office on iOS</span></span><br><span data-ttu-id="df454-454">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="df454-462">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-463">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-463">Office on Mac</span></span><br><span data-ttu-id="df454-464">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="df454-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="df454-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="df454-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="df454-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="df454-478">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-479">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-479">Office 2019 on Mac</span></span><br><span data-ttu-id="df454-480">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="df454-492">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-493">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-493">Office 2016 on Mac</span></span><br><span data-ttu-id="df454-494">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="df454-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="df454-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="df454-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="df454-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="df454-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="df454-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="df454-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="df454-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="df454-506">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-507">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="df454-507">Office on Android</span></span><br><span data-ttu-id="df454-508">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="df454-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="df454-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)</span><span class="sxs-lookup"><span data-stu-id="df454-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="df454-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="df454-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="df454-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="df454-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="df454-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="df454-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="df454-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="df454-517">Non disponible</span><span class="sxs-lookup"><span data-stu-id="df454-517">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="df454-518">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="df454-518">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="df454-519">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="df454-519">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="df454-520">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="df454-520">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="df454-521">Word</span><span class="sxs-lookup"><span data-stu-id="df454-521">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="df454-522">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="df454-522">Platform</span></span></th>
    <th><span data-ttu-id="df454-523">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-523">Extension points</span></span></th>
    <th><span data-ttu-id="df454-524">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-524">API requirement sets</span></span></th>
    <th><span data-ttu-id="df454-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-526">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-526">Office on the web</span></span></td>
    <td> <span data-ttu-id="df454-527">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-527">- TaskPane</span></span><br><span data-ttu-id="df454-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="df454-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-535">- BindingEvents</span></span><br><span data-ttu-id="df454-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-537">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-538">
         - File</span></span><br><span data-ttu-id="df454-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-540">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-543">
         - PdfFile</span></span><br><span data-ttu-id="df454-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-544">
         - Selection</span></span><br><span data-ttu-id="df454-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-545">
         - Settings</span></span><br><span data-ttu-id="df454-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-546">
         - TableBindings</span></span><br><span data-ttu-id="df454-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-547">
         - TableCoercion</span></span><br><span data-ttu-id="df454-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-548">
         - TextBindings</span></span><br><span data-ttu-id="df454-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-549">
         - TextCoercion</span></span><br><span data-ttu-id="df454-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-550">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-551">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="df454-551">Office on Windows</span></span><br><span data-ttu-id="df454-552">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-552">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-553">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-553">- TaskPane</span></span><br><span data-ttu-id="df454-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="df454-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-561">- BindingEvents</span></span><br><span data-ttu-id="df454-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-562">
         - CompressedFile</span></span><br><span data-ttu-id="df454-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-564">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-565">
         - File</span></span><br><span data-ttu-id="df454-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-567">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-570">
         - PdfFile</span></span><br><span data-ttu-id="df454-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-571">
         - Selection</span></span><br><span data-ttu-id="df454-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-572">
         - Settings</span></span><br><span data-ttu-id="df454-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-573">
         - TableBindings</span></span><br><span data-ttu-id="df454-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-574">
         - TableCoercion</span></span><br><span data-ttu-id="df454-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-575">
         - TextBindings</span></span><br><span data-ttu-id="df454-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-576">
         - TextCoercion</span></span><br><span data-ttu-id="df454-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-578">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-578">Office 2019 on Windows</span></span><br><span data-ttu-id="df454-579">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-580">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="df454-580">- TaskPane</span></span><br><span data-ttu-id="df454-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-587">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-587">- BindingEvents</span></span><br><span data-ttu-id="df454-588">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-588">
         - CompressedFile</span></span><br><span data-ttu-id="df454-589">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-589">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-590">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-590">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-591">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-591">
         - File</span></span><br><span data-ttu-id="df454-592">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-592">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-593">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-593">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-594">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-594">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-595">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-595">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-596">
         - PdfFile</span></span><br><span data-ttu-id="df454-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-597">
         - Selection</span></span><br><span data-ttu-id="df454-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-598">
         - Settings</span></span><br><span data-ttu-id="df454-599">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-599">
         - TableBindings</span></span><br><span data-ttu-id="df454-600">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-600">
         - TableCoercion</span></span><br><span data-ttu-id="df454-601">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-601">
         - TextBindings</span></span><br><span data-ttu-id="df454-602">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-602">
         - TextCoercion</span></span><br><span data-ttu-id="df454-603">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-603">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-604">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-604">Office 2016 on Windows</span></span><br><span data-ttu-id="df454-605">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-605">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-606">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-606">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="df454-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-610">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-610">- BindingEvents</span></span><br><span data-ttu-id="df454-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-611">
         - CompressedFile</span></span><br><span data-ttu-id="df454-612">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-612">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-613">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-613">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-614">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-614">
         - File</span></span><br><span data-ttu-id="df454-615">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-615">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-616">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-616">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-617">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-617">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-618">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-618">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-619">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-619">
         - PdfFile</span></span><br><span data-ttu-id="df454-620">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-620">
         - Selection</span></span><br><span data-ttu-id="df454-621">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-621">
         - Settings</span></span><br><span data-ttu-id="df454-622">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-622">
         - TableBindings</span></span><br><span data-ttu-id="df454-623">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-623">
         - TableCoercion</span></span><br><span data-ttu-id="df454-624">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-624">
         - TextBindings</span></span><br><span data-ttu-id="df454-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-625">
         - TextCoercion</span></span><br><span data-ttu-id="df454-626">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-626">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-627">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-627">Office 2013 on Windows</span></span><br><span data-ttu-id="df454-628">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-628">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-629">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-629">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="df454-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="df454-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-632">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-632">- BindingEvents</span></span><br><span data-ttu-id="df454-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-633">
         - CompressedFile</span></span><br><span data-ttu-id="df454-634">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-634">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-635">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-636">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-636">
         - File</span></span><br><span data-ttu-id="df454-637">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-637">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-638">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-638">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-639">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-639">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-640">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-640">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-641">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-641">
         - PdfFile</span></span><br><span data-ttu-id="df454-642">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-642">
         - Selection</span></span><br><span data-ttu-id="df454-643">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-643">
         - Settings</span></span><br><span data-ttu-id="df454-644">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-644">
         - TableBindings</span></span><br><span data-ttu-id="df454-645">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-645">
         - TableCoercion</span></span><br><span data-ttu-id="df454-646">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-646">
         - TextBindings</span></span><br><span data-ttu-id="df454-647">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-647">
         - TextCoercion</span></span><br><span data-ttu-id="df454-648">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-648">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-649">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="df454-649">Office on iPad</span></span><br><span data-ttu-id="df454-650">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-650">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-651">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="df454-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-657">- BindingEvents</span></span><br><span data-ttu-id="df454-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-658">
         - CompressedFile</span></span><br><span data-ttu-id="df454-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-660">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-661">
         - File</span></span><br><span data-ttu-id="df454-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-663">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-666">
         - PdfFile</span></span><br><span data-ttu-id="df454-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-667">
         - Selection</span></span><br><span data-ttu-id="df454-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-668">
         - Settings</span></span><br><span data-ttu-id="df454-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-669">
         - TableBindings</span></span><br><span data-ttu-id="df454-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-670">
         - TableCoercion</span></span><br><span data-ttu-id="df454-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-671">
         - TextBindings</span></span><br><span data-ttu-id="df454-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-672">
         - TextCoercion</span></span><br><span data-ttu-id="df454-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-674">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-674">Office on Mac</span></span><br><span data-ttu-id="df454-675">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-675">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-676">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-676">- TaskPane</span></span><br><span data-ttu-id="df454-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="df454-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-684">- BindingEvents</span></span><br><span data-ttu-id="df454-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-685">
         - CompressedFile</span></span><br><span data-ttu-id="df454-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-687">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-688">
         - File</span></span><br><span data-ttu-id="df454-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-690">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-693">
         - PdfFile</span></span><br><span data-ttu-id="df454-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-694">
         - Selection</span></span><br><span data-ttu-id="df454-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-695">
         - Settings</span></span><br><span data-ttu-id="df454-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-696">
         - TableBindings</span></span><br><span data-ttu-id="df454-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-697">
         - TableCoercion</span></span><br><span data-ttu-id="df454-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-698">
         - TextBindings</span></span><br><span data-ttu-id="df454-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-699">
         - TextCoercion</span></span><br><span data-ttu-id="df454-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-701">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-701">Office 2019 on Mac</span></span><br><span data-ttu-id="df454-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-703">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="df454-703">- TaskPane</span></span><br><span data-ttu-id="df454-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="df454-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="df454-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="df454-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="df454-710">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-710">- BindingEvents</span></span><br><span data-ttu-id="df454-711">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-711">
         - CompressedFile</span></span><br><span data-ttu-id="df454-712">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-712">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-713">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-713">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-714">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-714">
         - File</span></span><br><span data-ttu-id="df454-715">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-715">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-716">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-716">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-717">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-717">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-718">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-718">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-719">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-719">
         - PdfFile</span></span><br><span data-ttu-id="df454-720">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-720">
         - Selection</span></span><br><span data-ttu-id="df454-721">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-721">
         - Settings</span></span><br><span data-ttu-id="df454-722">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-722">
         - TableBindings</span></span><br><span data-ttu-id="df454-723">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-723">
         - TableCoercion</span></span><br><span data-ttu-id="df454-724">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-724">
         - TextBindings</span></span><br><span data-ttu-id="df454-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-725">
         - TextCoercion</span></span><br><span data-ttu-id="df454-726">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-726">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-727">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-727">Office 2016 on Mac</span></span><br><span data-ttu-id="df454-728">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-728">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-729">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-729">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="df454-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="df454-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="df454-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-733">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="df454-733">- BindingEvents</span></span><br><span data-ttu-id="df454-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-734">
         - CompressedFile</span></span><br><span data-ttu-id="df454-735">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="df454-735">
         - CustomXmlParts</span></span><br><span data-ttu-id="df454-736">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-736">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-737">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-737">
         - File</span></span><br><span data-ttu-id="df454-738">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-738">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-739">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="df454-739">
         - MatrixBindings</span></span><br><span data-ttu-id="df454-740">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-740">
         - MatrixCoercion</span></span><br><span data-ttu-id="df454-741">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-741">
         - OoxmlCoercion</span></span><br><span data-ttu-id="df454-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-742">
         - PdfFile</span></span><br><span data-ttu-id="df454-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-743">
         - Selection</span></span><br><span data-ttu-id="df454-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-744">
         - Settings</span></span><br><span data-ttu-id="df454-745">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="df454-745">
         - TableBindings</span></span><br><span data-ttu-id="df454-746">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-746">
         - TableCoercion</span></span><br><span data-ttu-id="df454-747">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="df454-747">
         - TextBindings</span></span><br><span data-ttu-id="df454-748">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-748">
         - TextCoercion</span></span><br><span data-ttu-id="df454-749">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="df454-749">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="df454-750">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="df454-750">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="df454-751">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="df454-751">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="df454-752">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="df454-752">Platform</span></span></th>
    <th><span data-ttu-id="df454-753">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-753">Extension points</span></span></th>
    <th><span data-ttu-id="df454-754">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-754">API requirement sets</span></span></th>
    <th><span data-ttu-id="df454-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-756">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-756">Office on the web</span></span></td>
    <td> <span data-ttu-id="df454-757">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-757">- Content</span></span><br><span data-ttu-id="df454-758">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-758">
         - TaskPane</span></span><br><span data-ttu-id="df454-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="df454-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="df454-764">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-764">- ActiveView</span></span><br><span data-ttu-id="df454-765">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-765">
         - CompressedFile</span></span><br><span data-ttu-id="df454-766">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-766">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-767">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-767">
         - File</span></span><br><span data-ttu-id="df454-768">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-768">
         - PdfFile</span></span><br><span data-ttu-id="df454-769">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-769">
         - Selection</span></span><br><span data-ttu-id="df454-770">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-770">
         - Settings</span></span><br><span data-ttu-id="df454-771">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-771">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-772">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="df454-772">Office on Windows</span></span><br><span data-ttu-id="df454-773">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-773">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-774">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-774">- Content</span></span><br><span data-ttu-id="df454-775">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-775">
         - TaskPane</span></span><br><span data-ttu-id="df454-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="df454-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="df454-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-781">- ActiveView</span></span><br><span data-ttu-id="df454-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-782">
         - CompressedFile</span></span><br><span data-ttu-id="df454-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-783">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-784">
         - File</span></span><br><span data-ttu-id="df454-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-785">
         - PdfFile</span></span><br><span data-ttu-id="df454-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-786">
         - Selection</span></span><br><span data-ttu-id="df454-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-787">
         - Settings</span></span><br><span data-ttu-id="df454-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-789">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-789">Office 2019 on Windows</span></span><br><span data-ttu-id="df454-790">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-791">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-791">- Content</span></span><br><span data-ttu-id="df454-792">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-792">
         - TaskPane</span></span><br><span data-ttu-id="df454-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-796">- ActiveView</span></span><br><span data-ttu-id="df454-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-797">
         - CompressedFile</span></span><br><span data-ttu-id="df454-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-798">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-799">
         - File</span></span><br><span data-ttu-id="df454-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-800">
         - PdfFile</span></span><br><span data-ttu-id="df454-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-801">
         - Selection</span></span><br><span data-ttu-id="df454-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-802">
         - Settings</span></span><br><span data-ttu-id="df454-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-804">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-804">Office 2016 on Windows</span></span><br><span data-ttu-id="df454-805">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-805">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-806">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-806">- Content</span></span><br><span data-ttu-id="df454-807">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="df454-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="df454-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-810">- ActiveView</span></span><br><span data-ttu-id="df454-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-811">
         - CompressedFile</span></span><br><span data-ttu-id="df454-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-812">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-813">
         - File</span></span><br><span data-ttu-id="df454-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-814">
         - PdfFile</span></span><br><span data-ttu-id="df454-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-815">
         - Selection</span></span><br><span data-ttu-id="df454-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-816">
         - Settings</span></span><br><span data-ttu-id="df454-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-818">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-818">Office 2013 on Windows</span></span><br><span data-ttu-id="df454-819">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-819">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-820">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-820">- Content</span></span><br><span data-ttu-id="df454-821">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-821">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="df454-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="df454-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="df454-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-824">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-824">- ActiveView</span></span><br><span data-ttu-id="df454-825">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-825">
         - CompressedFile</span></span><br><span data-ttu-id="df454-826">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-826">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-827">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-827">
         - File</span></span><br><span data-ttu-id="df454-828">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-828">
         - PdfFile</span></span><br><span data-ttu-id="df454-829">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-829">
         - Selection</span></span><br><span data-ttu-id="df454-830">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-830">
         - Settings</span></span><br><span data-ttu-id="df454-831">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-831">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-832">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="df454-832">Office on iPad</span></span><br><span data-ttu-id="df454-833">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-833">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-834">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-834">- Content</span></span><br><span data-ttu-id="df454-835">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-835">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="df454-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-839">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-839">- ActiveView</span></span><br><span data-ttu-id="df454-840">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-840">
         - CompressedFile</span></span><br><span data-ttu-id="df454-841">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-841">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-842">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-842">
         - File</span></span><br><span data-ttu-id="df454-843">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-843">
         - PdfFile</span></span><br><span data-ttu-id="df454-844">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-844">
         - Selection</span></span><br><span data-ttu-id="df454-845">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-845">
         - Settings</span></span><br><span data-ttu-id="df454-846">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-846">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-847">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-847">Office on Mac</span></span><br><span data-ttu-id="df454-848">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="df454-848">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="df454-849">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-849">- Content</span></span><br><span data-ttu-id="df454-850">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-850">
         - TaskPane</span></span><br><span data-ttu-id="df454-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="df454-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="df454-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="df454-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="df454-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-856">- ActiveView</span></span><br><span data-ttu-id="df454-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-857">
         - CompressedFile</span></span><br><span data-ttu-id="df454-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-858">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-859">
         - File</span></span><br><span data-ttu-id="df454-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-860">
         - PdfFile</span></span><br><span data-ttu-id="df454-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-861">
         - Selection</span></span><br><span data-ttu-id="df454-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-862">
         - Settings</span></span><br><span data-ttu-id="df454-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-863">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-864">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-864">Office 2019 on Mac</span></span><br><span data-ttu-id="df454-865">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-866">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-866">- Content</span></span><br><span data-ttu-id="df454-867">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-867">
         - TaskPane</span></span><br><span data-ttu-id="df454-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-871">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-871">- ActiveView</span></span><br><span data-ttu-id="df454-872">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-872">
         - CompressedFile</span></span><br><span data-ttu-id="df454-873">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-873">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-874">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-874">
         - File</span></span><br><span data-ttu-id="df454-875">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-875">
         - PdfFile</span></span><br><span data-ttu-id="df454-876">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-876">
         - Selection</span></span><br><span data-ttu-id="df454-877">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-877">
         - Settings</span></span><br><span data-ttu-id="df454-878">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-878">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-879">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="df454-879">Office 2016 on Mac</span></span><br><span data-ttu-id="df454-880">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-880">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-881">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-881">- Content</span></span><br><span data-ttu-id="df454-882">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-882">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="df454-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="df454-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-885">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="df454-885">- ActiveView</span></span><br><span data-ttu-id="df454-886">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="df454-886">
         - CompressedFile</span></span><br><span data-ttu-id="df454-887">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-887">
         - DocumentEvents</span></span><br><span data-ttu-id="df454-888">
         - File</span><span class="sxs-lookup"><span data-stu-id="df454-888">
         - File</span></span><br><span data-ttu-id="df454-889">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="df454-889">
         - PdfFile</span></span><br><span data-ttu-id="df454-890">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="df454-890">
         - Selection</span></span><br><span data-ttu-id="df454-891">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-891">
         - Settings</span></span><br><span data-ttu-id="df454-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-892">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="df454-893">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="df454-893">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="df454-894">OneNote</span><span class="sxs-lookup"><span data-stu-id="df454-894">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="df454-895">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="df454-895">Platform</span></span></th>
    <th><span data-ttu-id="df454-896">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-896">Extension points</span></span></th>
    <th><span data-ttu-id="df454-897">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-897">API requirement sets</span></span></th>
    <th><span data-ttu-id="df454-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-899">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="df454-899">Office on the web</span></span></td>
    <td> <span data-ttu-id="df454-900">- Contenu</span><span class="sxs-lookup"><span data-stu-id="df454-900">- Content</span></span><br><span data-ttu-id="df454-901">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-901">
         - TaskPane</span></span><br><span data-ttu-id="df454-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="df454-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="df454-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="df454-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="df454-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-906">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="df454-906">- DocumentEvents</span></span><br><span data-ttu-id="df454-907">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-907">
         - HtmlCoercion</span></span><br><span data-ttu-id="df454-908">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="df454-908">
         - Settings</span></span><br><span data-ttu-id="df454-909">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-909">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="df454-910">Projet</span><span class="sxs-lookup"><span data-stu-id="df454-910">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="df454-911">Plateforme</span><span class="sxs-lookup"><span data-stu-id="df454-911">Platform</span></span></th>
    <th><span data-ttu-id="df454-912">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="df454-912">Extension points</span></span></th>
    <th><span data-ttu-id="df454-913">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-913">API requirement sets</span></span></th>
    <th><span data-ttu-id="df454-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="df454-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-915">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-915">Office 2019 on Windows</span></span><br><span data-ttu-id="df454-916">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-916">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-917">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-917">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-919">- Selection</span><span class="sxs-lookup"><span data-stu-id="df454-919">- Selection</span></span><br><span data-ttu-id="df454-920">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-920">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-921">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-921">Office 2016 on Windows</span></span><br><span data-ttu-id="df454-922">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-922">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-923">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-923">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-925">- Selection</span><span class="sxs-lookup"><span data-stu-id="df454-925">- Selection</span></span><br><span data-ttu-id="df454-926">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-926">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="df454-927">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="df454-927">Office 2013 on Windows</span></span><br><span data-ttu-id="df454-928">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="df454-928">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="df454-929">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="df454-929">- TaskPane</span></span></td>
    <td> <span data-ttu-id="df454-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="df454-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="df454-931">- Selection</span><span class="sxs-lookup"><span data-stu-id="df454-931">- Selection</span></span><br><span data-ttu-id="df454-932">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="df454-932">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="df454-933">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="df454-933">See also</span></span>

- [<span data-ttu-id="df454-934">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="df454-934">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="df454-935">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="df454-935">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="df454-936">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="df454-936">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="df454-937">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="df454-937">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="df454-938">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="df454-938">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="df454-939">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="df454-939">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="df454-940">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="df454-940">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="df454-941">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="df454-941">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="df454-942">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="df454-942">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="df454-943">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="df454-943">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="df454-944">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="df454-944">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="df454-945">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="df454-945">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)