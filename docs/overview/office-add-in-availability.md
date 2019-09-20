---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: a3c580f32ad7cd384309a9b53e55ea488a470a90
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053325"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="dff63-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="dff63-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="dff63-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="dff63-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="dff63-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="dff63-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="dff63-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="dff63-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="dff63-108">Excel</span><span class="sxs-lookup"><span data-stu-id="dff63-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dff63-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="dff63-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dff63-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dff63-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dff63-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="dff63-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-114">- TaskPane</span></span><br><span data-ttu-id="dff63-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-115">
        - Content</span></span><br><span data-ttu-id="dff63-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-116">
        - Custom Functions</span></span><br><span data-ttu-id="dff63-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="dff63-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dff63-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dff63-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dff63-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-128">
        - BindingEvents</span></span><br><span data-ttu-id="dff63-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-129">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-130">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-131">
        - File</span></span><br><span data-ttu-id="dff63-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-132">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-134">
        - Selection</span></span><br><span data-ttu-id="dff63-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-135">
        - Settings</span></span><br><span data-ttu-id="dff63-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-136">
        - TableBindings</span></span><br><span data-ttu-id="dff63-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-137">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-138">
        - TextBindings</span></span><br><span data-ttu-id="dff63-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-140">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-140">Office on Windows</span></span><br><span data-ttu-id="dff63-141">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-142">- TaskPane</span></span><br><span data-ttu-id="dff63-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-143">
        - Content</span></span><br><span data-ttu-id="dff63-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-144">
        - Custom Functions</span></span><br><span data-ttu-id="dff63-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="dff63-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dff63-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dff63-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dff63-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dff63-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-158">
        - BindingEvents</span></span><br><span data-ttu-id="dff63-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-159">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-160">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-161">
        - File</span></span><br><span data-ttu-id="dff63-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-162">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-164">
        - Selection</span></span><br><span data-ttu-id="dff63-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-165">
        - Settings</span></span><br><span data-ttu-id="dff63-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-166">
        - TableBindings</span></span><br><span data-ttu-id="dff63-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-167">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-168">
        - TextBindings</span></span><br><span data-ttu-id="dff63-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-170">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-170">Office 2019 on Windows</span></span><br><span data-ttu-id="dff63-171">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dff63-172">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-172">- TaskPane</span></span><br><span data-ttu-id="dff63-173">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-173">
        - Content</span></span><br><span data-ttu-id="dff63-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dff63-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-185">- BindingEvents</span></span><br><span data-ttu-id="dff63-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-186">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-187">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-188">
        - File</span></span><br><span data-ttu-id="dff63-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-189">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-191">
        - Selection</span></span><br><span data-ttu-id="dff63-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-192">
        - Settings</span></span><br><span data-ttu-id="dff63-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-193">
        - TableBindings</span></span><br><span data-ttu-id="dff63-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-194">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-195">
        - TextBindings</span></span><br><span data-ttu-id="dff63-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-197">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-197">Office 2016 on Windows</span></span><br><span data-ttu-id="dff63-198">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dff63-199">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-199">- TaskPane</span></span><br><span data-ttu-id="dff63-200">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-200">
        - Content</span></span></td>
    <td><span data-ttu-id="dff63-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dff63-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-204">- BindingEvents</span></span><br><span data-ttu-id="dff63-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-205">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-206">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-207">
        - File</span></span><br><span data-ttu-id="dff63-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-208">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-210">
        - Selection</span></span><br><span data-ttu-id="dff63-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-211">
        - Settings</span></span><br><span data-ttu-id="dff63-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-212">
        - TableBindings</span></span><br><span data-ttu-id="dff63-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-213">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-214">
        - TextBindings</span></span><br><span data-ttu-id="dff63-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-216">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-216">Office 2013 on Windows</span></span><br><span data-ttu-id="dff63-217">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dff63-218">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-218">
        - TaskPane</span></span><br><span data-ttu-id="dff63-219">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="dff63-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dff63-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dff63-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-222">
        - BindingEvents</span></span><br><span data-ttu-id="dff63-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-223">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-224">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-225">
        - File</span></span><br><span data-ttu-id="dff63-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-226">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-228">
        - Selection</span></span><br><span data-ttu-id="dff63-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-229">
        - Settings</span></span><br><span data-ttu-id="dff63-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-230">
        - TableBindings</span></span><br><span data-ttu-id="dff63-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-231">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-232">
        - TextBindings</span></span><br><span data-ttu-id="dff63-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-234">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="dff63-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dff63-235">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dff63-236">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-236">- TaskPane</span></span><br><span data-ttu-id="dff63-237">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-237">
        - Content</span></span></td>
    <td><span data-ttu-id="dff63-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dff63-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dff63-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-249">- BindingEvents</span></span><br><span data-ttu-id="dff63-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-250">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-251">
        - File</span></span><br><span data-ttu-id="dff63-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-252">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-254">
        - Selection</span></span><br><span data-ttu-id="dff63-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-255">
        - Settings</span></span><br><span data-ttu-id="dff63-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-256">
        - TableBindings</span></span><br><span data-ttu-id="dff63-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-257">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-258">
        - TextBindings</span></span><br><span data-ttu-id="dff63-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-260">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-260">Office apps on Mac</span></span><br><span data-ttu-id="dff63-261">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dff63-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-262">- TaskPane</span></span><br><span data-ttu-id="dff63-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-263">
        - Content</span></span><br><span data-ttu-id="dff63-264">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-264">
        - Custom Functions</span></span><br><span data-ttu-id="dff63-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dff63-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dff63-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dff63-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dff63-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-278">- BindingEvents</span></span><br><span data-ttu-id="dff63-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-279">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-280">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-281">
        - File</span></span><br><span data-ttu-id="dff63-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-282">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-284">
        - PdfFile</span></span><br><span data-ttu-id="dff63-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-285">
        - Selection</span></span><br><span data-ttu-id="dff63-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-286">
        - Settings</span></span><br><span data-ttu-id="dff63-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-287">
        - TableBindings</span></span><br><span data-ttu-id="dff63-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-288">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-289">
        - TextBindings</span></span><br><span data-ttu-id="dff63-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-291">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-291">Office 2019 for Mac</span></span><br><span data-ttu-id="dff63-292">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dff63-293">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-293">- TaskPane</span></span><br><span data-ttu-id="dff63-294">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-294">
        - Content</span></span><br><span data-ttu-id="dff63-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dff63-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dff63-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dff63-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dff63-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dff63-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dff63-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dff63-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dff63-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dff63-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-306">- BindingEvents</span></span><br><span data-ttu-id="dff63-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-307">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-308">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-309">
        - File</span></span><br><span data-ttu-id="dff63-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-310">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-312">
        - PdfFile</span></span><br><span data-ttu-id="dff63-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-313">
        - Selection</span></span><br><span data-ttu-id="dff63-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-314">
        - Settings</span></span><br><span data-ttu-id="dff63-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-315">
        - TableBindings</span></span><br><span data-ttu-id="dff63-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-316">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-317">
        - TextBindings</span></span><br><span data-ttu-id="dff63-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-319">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dff63-320">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dff63-321">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-321">- TaskPane</span></span><br><span data-ttu-id="dff63-322">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-322">
        - Content</span></span></td>
    <td><span data-ttu-id="dff63-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dff63-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dff63-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dff63-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-326">- BindingEvents</span></span><br><span data-ttu-id="dff63-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-327">
        - CompressedFile</span></span><br><span data-ttu-id="dff63-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-328">
        - DocumentEvents</span></span><br><span data-ttu-id="dff63-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="dff63-329">
        - File</span></span><br><span data-ttu-id="dff63-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-330">
        - MatrixBindings</span></span><br><span data-ttu-id="dff63-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="dff63-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-332">
        - PdfFile</span></span><br><span data-ttu-id="dff63-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-333">
        - Selection</span></span><br><span data-ttu-id="dff63-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-334">
        - Settings</span></span><br><span data-ttu-id="dff63-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-335">
        - TableBindings</span></span><br><span data-ttu-id="dff63-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-336">
        - TableCoercion</span></span><br><span data-ttu-id="dff63-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-337">
        - TextBindings</span></span><br><span data-ttu-id="dff63-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dff63-339">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="dff63-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="dff63-340">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dff63-341">Plateforme</span><span class="sxs-lookup"><span data-stu-id="dff63-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dff63-342">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dff63-343">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dff63-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-345">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-345">Office on the web</span></span></td>
    <td><span data-ttu-id="dff63-346">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dff63-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-348">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-348">Office on Windows</span></span><br><span data-ttu-id="dff63-349">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dff63-350">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dff63-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-352">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-352">Office for Mac</span></span><br><span data-ttu-id="dff63-353">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="dff63-354">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dff63-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dff63-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="dff63-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="dff63-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dff63-357">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="dff63-357">Platform</span></span></th>
    <th><span data-ttu-id="dff63-358">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-358">Extension points</span></span></th>
    <th><span data-ttu-id="dff63-359">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="dff63-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-361">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-361">Office on the web</span></span><br><span data-ttu-id="dff63-362">(moderne)</span><span class="sxs-lookup"><span data-stu-id="dff63-362">Modern</span></span></td>
    <td> <span data-ttu-id="dff63-363">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-363">- Mail Read</span></span><br><span data-ttu-id="dff63-364">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-364">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dff63-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dff63-373">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-374">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-374">Office on the web</span></span><br><span data-ttu-id="dff63-375">(classique)</span><span class="sxs-lookup"><span data-stu-id="dff63-375">Classic.</span></span></td>
    <td> <span data-ttu-id="dff63-376">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-376">- Mail Read</span></span><br><span data-ttu-id="dff63-377">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-377">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dff63-385">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-386">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-386">Office on Windows</span></span><br><span data-ttu-id="dff63-387">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-388">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-388">- Mail Read</span></span><br><span data-ttu-id="dff63-389">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-389">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dff63-391">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="dff63-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dff63-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dff63-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dff63-399">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-400">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-400">Office 2019 on Windows</span></span><br><span data-ttu-id="dff63-401">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-402">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-402">- Mail Read</span></span><br><span data-ttu-id="dff63-403">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-403">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dff63-405">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="dff63-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dff63-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dff63-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dff63-413">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-414">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-414">Office 2016 on Windows</span></span><br><span data-ttu-id="dff63-415">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-416">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-416">- Mail Read</span></span><br><span data-ttu-id="dff63-417">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-417">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dff63-419">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="dff63-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dff63-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dff63-424">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-425">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-425">Office 2013 on Windows</span></span><br><span data-ttu-id="dff63-426">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-427">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-427">- Mail Read</span></span><br><span data-ttu-id="dff63-428">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="dff63-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="dff63-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dff63-433">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-434">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="dff63-434">Office apps on iOS</span></span><br><span data-ttu-id="dff63-435">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-436">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-436">- Mail Read</span></span><br><span data-ttu-id="dff63-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dff63-443">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-444">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-444">Office apps on Mac</span></span><br><span data-ttu-id="dff63-445">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-446">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-446">- Mail Read</span></span><br><span data-ttu-id="dff63-447">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-447">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dff63-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dff63-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dff63-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-457">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-457">Office 2019 for Mac</span></span><br><span data-ttu-id="dff63-458">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-459">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-459">- Mail Read</span></span><br><span data-ttu-id="dff63-460">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-460">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dff63-468">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-469">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dff63-470">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-471">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-471">- Mail Read</span></span><br><span data-ttu-id="dff63-472">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="dff63-472">
      - Mail Compose</span></span><br><span data-ttu-id="dff63-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dff63-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dff63-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dff63-480">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-481">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="dff63-481">Office apps on Android</span></span><br><span data-ttu-id="dff63-482">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-483">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="dff63-483">- Mail Read</span></span><br><span data-ttu-id="dff63-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dff63-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dff63-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dff63-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dff63-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dff63-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dff63-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dff63-490">Non disponible</span><span class="sxs-lookup"><span data-stu-id="dff63-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="dff63-491">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="dff63-491">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="dff63-492">Word</span><span class="sxs-lookup"><span data-stu-id="dff63-492">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dff63-493">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="dff63-493">Platform</span></span></th>
    <th><span data-ttu-id="dff63-494">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-494">Extension points</span></span></th>
    <th><span data-ttu-id="dff63-495">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-495">API requirement sets</span></span></th>
    <th><span data-ttu-id="dff63-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-497">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-497">Office on the web</span></span></td>
    <td> <span data-ttu-id="dff63-498">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-498">- TaskPane</span></span><br><span data-ttu-id="dff63-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dff63-506">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-506">- BindingEvents</span></span><br><span data-ttu-id="dff63-507">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-507">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-508">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-508">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-509">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-509">
         - File</span></span><br><span data-ttu-id="dff63-510">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-510">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-511">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-511">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-512">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-512">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-513">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-513">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-514">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-514">
         - PdfFile</span></span><br><span data-ttu-id="dff63-515">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-515">
         - Selection</span></span><br><span data-ttu-id="dff63-516">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-516">
         - Settings</span></span><br><span data-ttu-id="dff63-517">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-517">
         - TableBindings</span></span><br><span data-ttu-id="dff63-518">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-518">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-519">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-519">
         - TextBindings</span></span><br><span data-ttu-id="dff63-520">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-520">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-521">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-521">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-522">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-522">Office on Windows</span></span><br><span data-ttu-id="dff63-523">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-523">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-524">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-524">- TaskPane</span></span><br><span data-ttu-id="dff63-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dff63-532">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-532">- BindingEvents</span></span><br><span data-ttu-id="dff63-533">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-533">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-534">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-534">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-535">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-535">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-536">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-536">
         - File</span></span><br><span data-ttu-id="dff63-537">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-537">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-538">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-538">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-539">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-539">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-540">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-540">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-541">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-541">
         - PdfFile</span></span><br><span data-ttu-id="dff63-542">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-542">
         - Selection</span></span><br><span data-ttu-id="dff63-543">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-543">
         - Settings</span></span><br><span data-ttu-id="dff63-544">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-544">
         - TableBindings</span></span><br><span data-ttu-id="dff63-545">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-545">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-546">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-546">
         - TextBindings</span></span><br><span data-ttu-id="dff63-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-547">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-548">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-548">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-549">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-549">Office 2019 on Windows</span></span><br><span data-ttu-id="dff63-550">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-550">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-551">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="dff63-551">- TaskPane</span></span><br><span data-ttu-id="dff63-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-558">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-558">- BindingEvents</span></span><br><span data-ttu-id="dff63-559">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-559">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-560">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-560">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-561">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-561">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-562">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-562">
         - File</span></span><br><span data-ttu-id="dff63-563">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-563">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-564">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-564">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-565">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-565">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-566">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-566">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-567">
         - PdfFile</span></span><br><span data-ttu-id="dff63-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-568">
         - Selection</span></span><br><span data-ttu-id="dff63-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-569">
         - Settings</span></span><br><span data-ttu-id="dff63-570">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-570">
         - TableBindings</span></span><br><span data-ttu-id="dff63-571">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-571">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-572">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-572">
         - TextBindings</span></span><br><span data-ttu-id="dff63-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-573">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-574">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-574">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-575">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-575">Office 2016 on Windows</span></span><br><span data-ttu-id="dff63-576">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-576">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-577">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-577">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dff63-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-581">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-581">- BindingEvents</span></span><br><span data-ttu-id="dff63-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-582">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-583">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-583">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-584">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-584">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-585">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-585">
         - File</span></span><br><span data-ttu-id="dff63-586">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-586">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-587">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-587">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-588">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-588">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-589">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-589">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-590">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-590">
         - PdfFile</span></span><br><span data-ttu-id="dff63-591">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-591">
         - Selection</span></span><br><span data-ttu-id="dff63-592">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-592">
         - Settings</span></span><br><span data-ttu-id="dff63-593">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-593">
         - TableBindings</span></span><br><span data-ttu-id="dff63-594">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-594">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-595">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-595">
         - TextBindings</span></span><br><span data-ttu-id="dff63-596">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-596">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-597">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-597">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-598">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-598">Office 2013 on Windows</span></span><br><span data-ttu-id="dff63-599">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-599">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-600">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-600">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dff63-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dff63-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-603">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-603">- BindingEvents</span></span><br><span data-ttu-id="dff63-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-604">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-605">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-605">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-606">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-607">
         - File</span></span><br><span data-ttu-id="dff63-608">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-608">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-609">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-609">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-610">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-610">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-611">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-611">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-612">
         - PdfFile</span></span><br><span data-ttu-id="dff63-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-613">
         - Selection</span></span><br><span data-ttu-id="dff63-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-614">
         - Settings</span></span><br><span data-ttu-id="dff63-615">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-615">
         - TableBindings</span></span><br><span data-ttu-id="dff63-616">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-616">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-617">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-617">
         - TextBindings</span></span><br><span data-ttu-id="dff63-618">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-618">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-619">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-619">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-620">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="dff63-620">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dff63-621">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-621">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-622">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-622">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dff63-628">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-628">- BindingEvents</span></span><br><span data-ttu-id="dff63-629">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-629">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-630">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-630">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-631">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-631">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-632">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-632">
         - File</span></span><br><span data-ttu-id="dff63-633">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-633">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-634">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-634">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-635">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-635">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-636">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-636">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-637">
         - PdfFile</span></span><br><span data-ttu-id="dff63-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-638">
         - Selection</span></span><br><span data-ttu-id="dff63-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-639">
         - Settings</span></span><br><span data-ttu-id="dff63-640">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-640">
         - TableBindings</span></span><br><span data-ttu-id="dff63-641">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-641">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-642">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-642">
         - TextBindings</span></span><br><span data-ttu-id="dff63-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-643">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-644">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-644">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-645">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-645">Office apps on Mac</span></span><br><span data-ttu-id="dff63-646">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-646">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-647">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-647">- TaskPane</span></span><br><span data-ttu-id="dff63-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="dff63-655">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-655">- BindingEvents</span></span><br><span data-ttu-id="dff63-656">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-656">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-657">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-657">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-658">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-658">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-659">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-659">
         - File</span></span><br><span data-ttu-id="dff63-660">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-660">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-661">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-661">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-662">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-662">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-663">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-663">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-664">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-664">
         - PdfFile</span></span><br><span data-ttu-id="dff63-665">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-665">
         - Selection</span></span><br><span data-ttu-id="dff63-666">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-666">
         - Settings</span></span><br><span data-ttu-id="dff63-667">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-667">
         - TableBindings</span></span><br><span data-ttu-id="dff63-668">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-668">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-669">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-669">
         - TextBindings</span></span><br><span data-ttu-id="dff63-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-670">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-671">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-671">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-672">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-672">Office 2019 for Mac</span></span><br><span data-ttu-id="dff63-673">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-673">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-674">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="dff63-674">- TaskPane</span></span><br><span data-ttu-id="dff63-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dff63-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dff63-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dff63-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dff63-681">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-681">- BindingEvents</span></span><br><span data-ttu-id="dff63-682">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-682">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-683">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-683">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-684">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-684">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-685">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-685">
         - File</span></span><br><span data-ttu-id="dff63-686">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-686">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-687">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-687">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-688">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-688">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-689">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-689">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-690">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-690">
         - PdfFile</span></span><br><span data-ttu-id="dff63-691">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-691">
         - Selection</span></span><br><span data-ttu-id="dff63-692">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-692">
         - Settings</span></span><br><span data-ttu-id="dff63-693">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-693">
         - TableBindings</span></span><br><span data-ttu-id="dff63-694">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-694">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-695">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-695">
         - TextBindings</span></span><br><span data-ttu-id="dff63-696">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-696">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-697">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-697">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-698">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-698">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dff63-699">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-699">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-700">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-700">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dff63-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dff63-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dff63-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-704">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-704">- BindingEvents</span></span><br><span data-ttu-id="dff63-705">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-705">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-706">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dff63-706">
         - CustomXmlParts</span></span><br><span data-ttu-id="dff63-707">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-707">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-708">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-708">
         - File</span></span><br><span data-ttu-id="dff63-709">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-709">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-710">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-710">
         - MatrixBindings</span></span><br><span data-ttu-id="dff63-711">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-711">
         - MatrixCoercion</span></span><br><span data-ttu-id="dff63-712">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-712">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dff63-713">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-713">
         - PdfFile</span></span><br><span data-ttu-id="dff63-714">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-714">
         - Selection</span></span><br><span data-ttu-id="dff63-715">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-715">
         - Settings</span></span><br><span data-ttu-id="dff63-716">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-716">
         - TableBindings</span></span><br><span data-ttu-id="dff63-717">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-717">
         - TableCoercion</span></span><br><span data-ttu-id="dff63-718">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dff63-718">
         - TextBindings</span></span><br><span data-ttu-id="dff63-719">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-719">
         - TextCoercion</span></span><br><span data-ttu-id="dff63-720">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dff63-720">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="dff63-721">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="dff63-721">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="dff63-722">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dff63-722">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dff63-723">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="dff63-723">Platform</span></span></th>
    <th><span data-ttu-id="dff63-724">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-724">Extension points</span></span></th>
    <th><span data-ttu-id="dff63-725">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-725">API requirement sets</span></span></th>
    <th><span data-ttu-id="dff63-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-727">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-727">Office on the web</span></span></td>
    <td> <span data-ttu-id="dff63-728">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-728">- Content</span></span><br><span data-ttu-id="dff63-729">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-729">
         - TaskPane</span></span><br><span data-ttu-id="dff63-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dff63-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dff63-735">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-735">- ActiveView</span></span><br><span data-ttu-id="dff63-736">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-736">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-737">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-737">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-738">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-738">
         - File</span></span><br><span data-ttu-id="dff63-739">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-739">
         - PdfFile</span></span><br><span data-ttu-id="dff63-740">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-740">
         - Selection</span></span><br><span data-ttu-id="dff63-741">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-741">
         - Settings</span></span><br><span data-ttu-id="dff63-742">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-742">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-743">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-743">Office on Windows</span></span><br><span data-ttu-id="dff63-744">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-744">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-745">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-745">- Content</span></span><br><span data-ttu-id="dff63-746">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-746">
         - TaskPane</span></span><br><span data-ttu-id="dff63-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dff63-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dff63-752">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-752">- ActiveView</span></span><br><span data-ttu-id="dff63-753">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-753">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-754">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-754">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-755">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-755">
         - File</span></span><br><span data-ttu-id="dff63-756">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-756">
         - PdfFile</span></span><br><span data-ttu-id="dff63-757">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-757">
         - Selection</span></span><br><span data-ttu-id="dff63-758">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-758">
         - Settings</span></span><br><span data-ttu-id="dff63-759">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-759">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-760">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-760">Office 2019 on Windows</span></span><br><span data-ttu-id="dff63-761">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-761">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-762">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-762">- Content</span></span><br><span data-ttu-id="dff63-763">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-763">
         - TaskPane</span></span><br><span data-ttu-id="dff63-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-767">- ActiveView</span></span><br><span data-ttu-id="dff63-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-768">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-769">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-770">
         - File</span></span><br><span data-ttu-id="dff63-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-771">
         - PdfFile</span></span><br><span data-ttu-id="dff63-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-772">
         - Selection</span></span><br><span data-ttu-id="dff63-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-773">
         - Settings</span></span><br><span data-ttu-id="dff63-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-775">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-775">Office 2016 on Windows</span></span><br><span data-ttu-id="dff63-776">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-776">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-777">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-777">- Content</span></span><br><span data-ttu-id="dff63-778">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-778">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dff63-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dff63-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-781">- ActiveView</span></span><br><span data-ttu-id="dff63-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-782">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-783">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-784">
         - File</span></span><br><span data-ttu-id="dff63-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-785">
         - PdfFile</span></span><br><span data-ttu-id="dff63-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-786">
         - Selection</span></span><br><span data-ttu-id="dff63-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-787">
         - Settings</span></span><br><span data-ttu-id="dff63-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-789">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-789">Office 2013 on Windows</span></span><br><span data-ttu-id="dff63-790">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-791">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-791">- Content</span></span><br><span data-ttu-id="dff63-792">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-792">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="dff63-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dff63-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dff63-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-795">- ActiveView</span></span><br><span data-ttu-id="dff63-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-796">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-797">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-798">
         - File</span></span><br><span data-ttu-id="dff63-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-799">
         - PdfFile</span></span><br><span data-ttu-id="dff63-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-800">
         - Selection</span></span><br><span data-ttu-id="dff63-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-801">
         - Settings</span></span><br><span data-ttu-id="dff63-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-803">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="dff63-803">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dff63-804">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-804">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-805">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-805">- Content</span></span><br><span data-ttu-id="dff63-806">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dff63-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-810">- ActiveView</span></span><br><span data-ttu-id="dff63-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-811">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-812">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-813">
         - File</span></span><br><span data-ttu-id="dff63-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-814">
         - PdfFile</span></span><br><span data-ttu-id="dff63-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-815">
         - Selection</span></span><br><span data-ttu-id="dff63-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-816">
         - Settings</span></span><br><span data-ttu-id="dff63-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-818">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-818">Office apps on Mac</span></span><br><span data-ttu-id="dff63-819">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="dff63-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dff63-820">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-820">- Content</span></span><br><span data-ttu-id="dff63-821">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-821">
         - TaskPane</span></span><br><span data-ttu-id="dff63-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dff63-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dff63-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dff63-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dff63-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-827">- ActiveView</span></span><br><span data-ttu-id="dff63-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-828">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-829">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-830">
         - File</span></span><br><span data-ttu-id="dff63-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-831">
         - PdfFile</span></span><br><span data-ttu-id="dff63-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-832">
         - Selection</span></span><br><span data-ttu-id="dff63-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-833">
         - Settings</span></span><br><span data-ttu-id="dff63-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-835">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-835">Office 2019 for Mac</span></span><br><span data-ttu-id="dff63-836">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-836">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-837">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-837">- Content</span></span><br><span data-ttu-id="dff63-838">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-838">
         - TaskPane</span></span><br><span data-ttu-id="dff63-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-842">- ActiveView</span></span><br><span data-ttu-id="dff63-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-843">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-844">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-845">
         - File</span></span><br><span data-ttu-id="dff63-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-846">
         - PdfFile</span></span><br><span data-ttu-id="dff63-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-847">
         - Selection</span></span><br><span data-ttu-id="dff63-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-848">
         - Settings</span></span><br><span data-ttu-id="dff63-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-850">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-850">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dff63-851">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-851">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-852">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-852">- Content</span></span><br><span data-ttu-id="dff63-853">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-853">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dff63-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dff63-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dff63-856">- ActiveView</span></span><br><span data-ttu-id="dff63-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dff63-857">
         - CompressedFile</span></span><br><span data-ttu-id="dff63-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-858">
         - DocumentEvents</span></span><br><span data-ttu-id="dff63-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="dff63-859">
         - File</span></span><br><span data-ttu-id="dff63-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dff63-860">
         - PdfFile</span></span><br><span data-ttu-id="dff63-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-861">
         - Selection</span></span><br><span data-ttu-id="dff63-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-862">
         - Settings</span></span><br><span data-ttu-id="dff63-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-863">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dff63-864">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="dff63-864">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="dff63-865">OneNote</span><span class="sxs-lookup"><span data-stu-id="dff63-865">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dff63-866">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="dff63-866">Platform</span></span></th>
    <th><span data-ttu-id="dff63-867">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-867">Extension points</span></span></th>
    <th><span data-ttu-id="dff63-868">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-868">API requirement sets</span></span></th>
    <th><span data-ttu-id="dff63-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-870">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="dff63-870">Office on the web</span></span></td>
    <td> <span data-ttu-id="dff63-871">- Contenu</span><span class="sxs-lookup"><span data-stu-id="dff63-871">- Content</span></span><br><span data-ttu-id="dff63-872">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-872">
         - TaskPane</span></span><br><span data-ttu-id="dff63-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="dff63-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dff63-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="dff63-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dff63-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-877">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dff63-877">- DocumentEvents</span></span><br><span data-ttu-id="dff63-878">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-878">
         - HtmlCoercion</span></span><br><span data-ttu-id="dff63-879">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dff63-879">
         - Settings</span></span><br><span data-ttu-id="dff63-880">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-880">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="dff63-881">Projet</span><span class="sxs-lookup"><span data-stu-id="dff63-881">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dff63-882">Plateforme</span><span class="sxs-lookup"><span data-stu-id="dff63-882">Platform</span></span></th>
    <th><span data-ttu-id="dff63-883">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="dff63-883">Extension points</span></span></th>
    <th><span data-ttu-id="dff63-884">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="dff63-884">API requirement sets</span></span></th>
    <th><span data-ttu-id="dff63-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="dff63-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-886">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-886">Office 2019 on Windows</span></span><br><span data-ttu-id="dff63-887">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-888">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-890">- Selection</span></span><br><span data-ttu-id="dff63-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-891">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-892">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-892">Office 2016 on Windows</span></span><br><span data-ttu-id="dff63-893">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-893">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-894">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-894">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-896">- Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-896">- Selection</span></span><br><span data-ttu-id="dff63-897">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-897">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dff63-898">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="dff63-898">Office 2013 on Windows</span></span><br><span data-ttu-id="dff63-899">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="dff63-899">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dff63-900">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="dff63-900">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dff63-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dff63-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dff63-902">- Selection</span><span class="sxs-lookup"><span data-stu-id="dff63-902">- Selection</span></span><br><span data-ttu-id="dff63-903">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dff63-903">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="dff63-904">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dff63-904">See also</span></span>

- [<span data-ttu-id="dff63-905">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="dff63-905">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="dff63-906">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dff63-906">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="dff63-907">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="dff63-907">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="dff63-908">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="dff63-908">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="dff63-909">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="dff63-909">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="dff63-910">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="dff63-910">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="dff63-911">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="dff63-911">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="dff63-912">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="dff63-912">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="dff63-913">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="dff63-913">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="dff63-914">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="dff63-914">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="dff63-915">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="dff63-915">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
