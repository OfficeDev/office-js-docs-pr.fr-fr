---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554019"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="96776-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="96776-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="96776-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="96776-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="96776-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="96776-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="96776-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="96776-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="96776-108">Excel</span><span class="sxs-lookup"><span data-stu-id="96776-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="96776-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="96776-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="96776-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="96776-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="96776-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="96776-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-114">- TaskPane</span></span><br><span data-ttu-id="96776-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-115">
        - Content</span></span><br><span data-ttu-id="96776-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-116">
        - Custom Functions</span></span><br><span data-ttu-id="96776-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="96776-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="96776-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="96776-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="96776-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="96776-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="96776-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="96776-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="96776-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-130">
        - BindingEvents</span></span><br><span data-ttu-id="96776-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-131">
        - CompressedFile</span></span><br><span data-ttu-id="96776-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-132">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-133">
        - File</span></span><br><span data-ttu-id="96776-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-134">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-136">
        - Selection</span></span><br><span data-ttu-id="96776-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-137">
        - Settings</span></span><br><span data-ttu-id="96776-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-138">
        - TableBindings</span></span><br><span data-ttu-id="96776-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-139">
        - TableCoercion</span></span><br><span data-ttu-id="96776-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-140">
        - TextBindings</span></span><br><span data-ttu-id="96776-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="96776-142">Office on Windows</span></span><br><span data-ttu-id="96776-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-144">- TaskPane</span></span><br><span data-ttu-id="96776-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-145">
        - Content</span></span><br><span data-ttu-id="96776-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-146">
        - Custom Functions</span></span><br><span data-ttu-id="96776-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="96776-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="96776-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="96776-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="96776-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="96776-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="96776-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="96776-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-161">
        - BindingEvents</span></span><br><span data-ttu-id="96776-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-162">
        - CompressedFile</span></span><br><span data-ttu-id="96776-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-163">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-164">
        - File</span></span><br><span data-ttu-id="96776-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-165">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-167">
        - Selection</span></span><br><span data-ttu-id="96776-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-168">
        - Settings</span></span><br><span data-ttu-id="96776-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-169">
        - TableBindings</span></span><br><span data-ttu-id="96776-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-170">
        - TableCoercion</span></span><br><span data-ttu-id="96776-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-171">
        - TextBindings</span></span><br><span data-ttu-id="96776-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-173">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-173">Office 2019 on Windows</span></span><br><span data-ttu-id="96776-174">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="96776-175">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-175">- TaskPane</span></span><br><span data-ttu-id="96776-176">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-176">
        - Content</span></span><br><span data-ttu-id="96776-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="96776-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-188">- BindingEvents</span></span><br><span data-ttu-id="96776-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-189">
        - CompressedFile</span></span><br><span data-ttu-id="96776-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-190">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-191">
        - File</span></span><br><span data-ttu-id="96776-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-192">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-194">
        - Selection</span></span><br><span data-ttu-id="96776-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-195">
        - Settings</span></span><br><span data-ttu-id="96776-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-196">
        - TableBindings</span></span><br><span data-ttu-id="96776-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-197">
        - TableCoercion</span></span><br><span data-ttu-id="96776-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-198">
        - TextBindings</span></span><br><span data-ttu-id="96776-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-200">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-200">Office 2016 on Windows</span></span><br><span data-ttu-id="96776-201">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="96776-202">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-202">- TaskPane</span></span><br><span data-ttu-id="96776-203">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-203">
        - Content</span></span></td>
    <td><span data-ttu-id="96776-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="96776-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-207">- BindingEvents</span></span><br><span data-ttu-id="96776-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-208">
        - CompressedFile</span></span><br><span data-ttu-id="96776-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-209">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-210">
        - File</span></span><br><span data-ttu-id="96776-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-211">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-213">
        - Selection</span></span><br><span data-ttu-id="96776-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-214">
        - Settings</span></span><br><span data-ttu-id="96776-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-215">
        - TableBindings</span></span><br><span data-ttu-id="96776-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-216">
        - TableCoercion</span></span><br><span data-ttu-id="96776-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-217">
        - TextBindings</span></span><br><span data-ttu-id="96776-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-219">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-219">Office 2013 on Windows</span></span><br><span data-ttu-id="96776-220">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="96776-221">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-221">
        - TaskPane</span></span><br><span data-ttu-id="96776-222">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="96776-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="96776-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="96776-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-225">
        - BindingEvents</span></span><br><span data-ttu-id="96776-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-226">
        - CompressedFile</span></span><br><span data-ttu-id="96776-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-227">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-228">
        - File</span></span><br><span data-ttu-id="96776-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-229">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-231">
        - Selection</span></span><br><span data-ttu-id="96776-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-232">
        - Settings</span></span><br><span data-ttu-id="96776-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-233">
        - TableBindings</span></span><br><span data-ttu-id="96776-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-234">
        - TableCoercion</span></span><br><span data-ttu-id="96776-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-235">
        - TextBindings</span></span><br><span data-ttu-id="96776-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-237">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="96776-237">Office on iPad</span></span><br><span data-ttu-id="96776-238">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="96776-239">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-239">- TaskPane</span></span><br><span data-ttu-id="96776-240">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-240">
        - Content</span></span></td>
    <td><span data-ttu-id="96776-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="96776-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="96776-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="96776-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="96776-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-253">- BindingEvents</span></span><br><span data-ttu-id="96776-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-254">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-255">
        - File</span></span><br><span data-ttu-id="96776-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-256">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-258">
        - Selection</span></span><br><span data-ttu-id="96776-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-259">
        - Settings</span></span><br><span data-ttu-id="96776-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-260">
        - TableBindings</span></span><br><span data-ttu-id="96776-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-261">
        - TableCoercion</span></span><br><span data-ttu-id="96776-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-262">
        - TextBindings</span></span><br><span data-ttu-id="96776-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-264">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-264">Office on Mac</span></span><br><span data-ttu-id="96776-265">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="96776-266">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-266">- TaskPane</span></span><br><span data-ttu-id="96776-267">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-267">
        - Content</span></span><br><span data-ttu-id="96776-268">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-268">
        - Custom Functions</span></span><br><span data-ttu-id="96776-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="96776-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="96776-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="96776-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="96776-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="96776-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="96776-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-283">- BindingEvents</span></span><br><span data-ttu-id="96776-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-284">
        - CompressedFile</span></span><br><span data-ttu-id="96776-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-285">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-286">
        - File</span></span><br><span data-ttu-id="96776-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-287">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-289">
        - PdfFile</span></span><br><span data-ttu-id="96776-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-290">
        - Selection</span></span><br><span data-ttu-id="96776-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-291">
        - Settings</span></span><br><span data-ttu-id="96776-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-292">
        - TableBindings</span></span><br><span data-ttu-id="96776-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-293">
        - TableCoercion</span></span><br><span data-ttu-id="96776-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-294">
        - TextBindings</span></span><br><span data-ttu-id="96776-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-296">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-296">Office 2019 on Mac</span></span><br><span data-ttu-id="96776-297">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="96776-298">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-298">- TaskPane</span></span><br><span data-ttu-id="96776-299">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-299">
        - Content</span></span><br><span data-ttu-id="96776-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="96776-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="96776-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="96776-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="96776-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="96776-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="96776-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="96776-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="96776-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-311">- BindingEvents</span></span><br><span data-ttu-id="96776-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-312">
        - CompressedFile</span></span><br><span data-ttu-id="96776-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-313">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-314">
        - File</span></span><br><span data-ttu-id="96776-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-315">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-317">
        - PdfFile</span></span><br><span data-ttu-id="96776-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-318">
        - Selection</span></span><br><span data-ttu-id="96776-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-319">
        - Settings</span></span><br><span data-ttu-id="96776-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-320">
        - TableBindings</span></span><br><span data-ttu-id="96776-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-321">
        - TableCoercion</span></span><br><span data-ttu-id="96776-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-322">
        - TextBindings</span></span><br><span data-ttu-id="96776-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-324">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-324">Office 2016 on Mac</span></span><br><span data-ttu-id="96776-325">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="96776-326">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-326">- TaskPane</span></span><br><span data-ttu-id="96776-327">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-327">
        - Content</span></span></td>
    <td><span data-ttu-id="96776-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="96776-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="96776-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="96776-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-331">- BindingEvents</span></span><br><span data-ttu-id="96776-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-332">
        - CompressedFile</span></span><br><span data-ttu-id="96776-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-333">
        - DocumentEvents</span></span><br><span data-ttu-id="96776-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="96776-334">
        - File</span></span><br><span data-ttu-id="96776-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-335">
        - MatrixBindings</span></span><br><span data-ttu-id="96776-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="96776-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-337">
        - PdfFile</span></span><br><span data-ttu-id="96776-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-338">
        - Selection</span></span><br><span data-ttu-id="96776-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-339">
        - Settings</span></span><br><span data-ttu-id="96776-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-340">
        - TableBindings</span></span><br><span data-ttu-id="96776-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-341">
        - TableCoercion</span></span><br><span data-ttu-id="96776-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-342">
        - TextBindings</span></span><br><span data-ttu-id="96776-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="96776-344">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="96776-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="96776-345">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="96776-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="96776-346">Plateforme</span><span class="sxs-lookup"><span data-stu-id="96776-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="96776-347">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="96776-348">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="96776-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-350">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-350">Office on the web</span></span></td>
    <td><span data-ttu-id="96776-351">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="96776-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-353">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="96776-353">Office on Windows</span></span><br><span data-ttu-id="96776-354">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="96776-355">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="96776-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-357">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="96776-357">Office for Mac</span></span><br><span data-ttu-id="96776-358">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="96776-359">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="96776-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="96776-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="96776-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="96776-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="96776-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="96776-362">Platform</span></span></th>
    <th><span data-ttu-id="96776-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-363">Extension points</span></span></th>
    <th><span data-ttu-id="96776-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="96776-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-366">Office on the web</span></span><br><span data-ttu-id="96776-367">(moderne)</span><span class="sxs-lookup"><span data-stu-id="96776-367">(modern)</span></span></td>
    <td> <span data-ttu-id="96776-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-368">- Mail Read</span></span><br><span data-ttu-id="96776-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-369">
      - Mail Compose</span></span><br><span data-ttu-id="96776-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="96776-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="96776-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="96776-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-380">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-380">Office on the web</span></span><br><span data-ttu-id="96776-381">(classique)</span><span class="sxs-lookup"><span data-stu-id="96776-381">(classic)</span></span></td>
    <td> <span data-ttu-id="96776-382">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-382">- Mail Read</span></span><br><span data-ttu-id="96776-383">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-383">
      - Mail Compose</span></span><br><span data-ttu-id="96776-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="96776-391">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-392">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="96776-392">Office on Windows</span></span><br><span data-ttu-id="96776-393">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-394">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-394">- Mail Read</span></span><br><span data-ttu-id="96776-395">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-395">
      - Mail Compose</span></span><br><span data-ttu-id="96776-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="96776-397">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="96776-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="96776-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="96776-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="96776-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="96776-406">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-407">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-407">Office 2019 on Windows</span></span><br><span data-ttu-id="96776-408">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-409">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-409">- Mail Read</span></span><br><span data-ttu-id="96776-410">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-410">
      - Mail Compose</span></span><br><span data-ttu-id="96776-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="96776-412">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="96776-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="96776-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="96776-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="96776-420">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-421">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-421">Office 2016 on Windows</span></span><br><span data-ttu-id="96776-422">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-423">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-423">- Mail Read</span></span><br><span data-ttu-id="96776-424">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-424">
      - Mail Compose</span></span><br><span data-ttu-id="96776-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="96776-426">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="96776-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="96776-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="96776-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-432">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-432">Office 2013 on Windows</span></span><br><span data-ttu-id="96776-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-434">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-434">- Mail Read</span></span><br><span data-ttu-id="96776-435">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="96776-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="96776-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="96776-440">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-441">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="96776-441">Office on iOS</span></span><br><span data-ttu-id="96776-442">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-443">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-443">- Mail Read</span></span><br><span data-ttu-id="96776-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="96776-450">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-451">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-451">Office on Mac</span></span><br><span data-ttu-id="96776-452">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-453">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-453">- Mail Read</span></span><br><span data-ttu-id="96776-454">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-454">
      - Mail Compose</span></span><br><span data-ttu-id="96776-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="96776-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="96776-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="96776-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="96776-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="96776-464">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-465">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-465">Office 2019 on Mac</span></span><br><span data-ttu-id="96776-466">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-467">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-467">- Mail Read</span></span><br><span data-ttu-id="96776-468">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-468">
      - Mail Compose</span></span><br><span data-ttu-id="96776-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="96776-476">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-477">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-477">Office 2016 on Mac</span></span><br><span data-ttu-id="96776-478">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-479">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-479">- Mail Read</span></span><br><span data-ttu-id="96776-480">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="96776-480">
      - Mail Compose</span></span><br><span data-ttu-id="96776-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="96776-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="96776-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="96776-488">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-489">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="96776-489">Office on Android</span></span><br><span data-ttu-id="96776-490">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-491">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="96776-491">- Mail Read</span></span><br><span data-ttu-id="96776-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="96776-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="96776-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="96776-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="96776-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="96776-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="96776-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="96776-498">Non disponible</span><span class="sxs-lookup"><span data-stu-id="96776-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="96776-499">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="96776-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="96776-500">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="96776-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="96776-501">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="96776-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="96776-502">Word</span><span class="sxs-lookup"><span data-stu-id="96776-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="96776-503">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="96776-503">Platform</span></span></th>
    <th><span data-ttu-id="96776-504">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-504">Extension points</span></span></th>
    <th><span data-ttu-id="96776-505">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="96776-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-507">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="96776-508">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-508">- TaskPane</span></span><br><span data-ttu-id="96776-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="96776-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-516">- BindingEvents</span></span><br><span data-ttu-id="96776-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-518">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-519">
         - File</span></span><br><span data-ttu-id="96776-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-521">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-524">
         - PdfFile</span></span><br><span data-ttu-id="96776-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-525">
         - Selection</span></span><br><span data-ttu-id="96776-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-526">
         - Settings</span></span><br><span data-ttu-id="96776-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-527">
         - TableBindings</span></span><br><span data-ttu-id="96776-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-528">
         - TableCoercion</span></span><br><span data-ttu-id="96776-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-529">
         - TextBindings</span></span><br><span data-ttu-id="96776-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-530">
         - TextCoercion</span></span><br><span data-ttu-id="96776-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-532">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="96776-532">Office on Windows</span></span><br><span data-ttu-id="96776-533">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-534">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-534">- TaskPane</span></span><br><span data-ttu-id="96776-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="96776-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-542">- BindingEvents</span></span><br><span data-ttu-id="96776-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-543">
         - CompressedFile</span></span><br><span data-ttu-id="96776-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-545">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-546">
         - File</span></span><br><span data-ttu-id="96776-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-548">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-551">
         - PdfFile</span></span><br><span data-ttu-id="96776-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-552">
         - Selection</span></span><br><span data-ttu-id="96776-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-553">
         - Settings</span></span><br><span data-ttu-id="96776-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-554">
         - TableBindings</span></span><br><span data-ttu-id="96776-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-555">
         - TableCoercion</span></span><br><span data-ttu-id="96776-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-556">
         - TextBindings</span></span><br><span data-ttu-id="96776-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-557">
         - TextCoercion</span></span><br><span data-ttu-id="96776-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-559">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-559">Office 2019 on Windows</span></span><br><span data-ttu-id="96776-560">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-561">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="96776-561">- TaskPane</span></span><br><span data-ttu-id="96776-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-568">- BindingEvents</span></span><br><span data-ttu-id="96776-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-569">
         - CompressedFile</span></span><br><span data-ttu-id="96776-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-571">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-572">
         - File</span></span><br><span data-ttu-id="96776-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-574">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-577">
         - PdfFile</span></span><br><span data-ttu-id="96776-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-578">
         - Selection</span></span><br><span data-ttu-id="96776-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-579">
         - Settings</span></span><br><span data-ttu-id="96776-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-580">
         - TableBindings</span></span><br><span data-ttu-id="96776-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-581">
         - TableCoercion</span></span><br><span data-ttu-id="96776-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-582">
         - TextBindings</span></span><br><span data-ttu-id="96776-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-583">
         - TextCoercion</span></span><br><span data-ttu-id="96776-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-585">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-585">Office 2016 on Windows</span></span><br><span data-ttu-id="96776-586">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-587">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="96776-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-591">- BindingEvents</span></span><br><span data-ttu-id="96776-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-592">
         - CompressedFile</span></span><br><span data-ttu-id="96776-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-594">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-595">
         - File</span></span><br><span data-ttu-id="96776-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-597">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-600">
         - PdfFile</span></span><br><span data-ttu-id="96776-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-601">
         - Selection</span></span><br><span data-ttu-id="96776-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-602">
         - Settings</span></span><br><span data-ttu-id="96776-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-603">
         - TableBindings</span></span><br><span data-ttu-id="96776-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-604">
         - TableCoercion</span></span><br><span data-ttu-id="96776-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-605">
         - TextBindings</span></span><br><span data-ttu-id="96776-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-606">
         - TextCoercion</span></span><br><span data-ttu-id="96776-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-608">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-608">Office 2013 on Windows</span></span><br><span data-ttu-id="96776-609">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-610">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="96776-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="96776-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-613">- BindingEvents</span></span><br><span data-ttu-id="96776-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-614">
         - CompressedFile</span></span><br><span data-ttu-id="96776-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-616">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-617">
         - File</span></span><br><span data-ttu-id="96776-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-619">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-622">
         - PdfFile</span></span><br><span data-ttu-id="96776-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-623">
         - Selection</span></span><br><span data-ttu-id="96776-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-624">
         - Settings</span></span><br><span data-ttu-id="96776-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-625">
         - TableBindings</span></span><br><span data-ttu-id="96776-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-626">
         - TableCoercion</span></span><br><span data-ttu-id="96776-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-627">
         - TextBindings</span></span><br><span data-ttu-id="96776-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-628">
         - TextCoercion</span></span><br><span data-ttu-id="96776-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-630">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="96776-630">Office on iPad</span></span><br><span data-ttu-id="96776-631">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="96776-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-638">- BindingEvents</span></span><br><span data-ttu-id="96776-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-639">
         - CompressedFile</span></span><br><span data-ttu-id="96776-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-641">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-642">
         - File</span></span><br><span data-ttu-id="96776-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-644">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-647">
         - PdfFile</span></span><br><span data-ttu-id="96776-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-648">
         - Selection</span></span><br><span data-ttu-id="96776-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-649">
         - Settings</span></span><br><span data-ttu-id="96776-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-650">
         - TableBindings</span></span><br><span data-ttu-id="96776-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-651">
         - TableCoercion</span></span><br><span data-ttu-id="96776-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-652">
         - TextBindings</span></span><br><span data-ttu-id="96776-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-653">
         - TextCoercion</span></span><br><span data-ttu-id="96776-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-655">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-655">Office on Mac</span></span><br><span data-ttu-id="96776-656">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-657">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-657">- TaskPane</span></span><br><span data-ttu-id="96776-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="96776-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-665">- BindingEvents</span></span><br><span data-ttu-id="96776-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-666">
         - CompressedFile</span></span><br><span data-ttu-id="96776-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-668">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-669">
         - File</span></span><br><span data-ttu-id="96776-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-671">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-674">
         - PdfFile</span></span><br><span data-ttu-id="96776-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-675">
         - Selection</span></span><br><span data-ttu-id="96776-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-676">
         - Settings</span></span><br><span data-ttu-id="96776-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-677">
         - TableBindings</span></span><br><span data-ttu-id="96776-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-678">
         - TableCoercion</span></span><br><span data-ttu-id="96776-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-679">
         - TextBindings</span></span><br><span data-ttu-id="96776-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-680">
         - TextCoercion</span></span><br><span data-ttu-id="96776-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-682">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-682">Office 2019 on Mac</span></span><br><span data-ttu-id="96776-683">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-684">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="96776-684">- TaskPane</span></span><br><span data-ttu-id="96776-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="96776-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="96776-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="96776-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="96776-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-691">- BindingEvents</span></span><br><span data-ttu-id="96776-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-692">
         - CompressedFile</span></span><br><span data-ttu-id="96776-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-694">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-695">
         - File</span></span><br><span data-ttu-id="96776-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-697">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-700">
         - PdfFile</span></span><br><span data-ttu-id="96776-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-701">
         - Selection</span></span><br><span data-ttu-id="96776-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-702">
         - Settings</span></span><br><span data-ttu-id="96776-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-703">
         - TableBindings</span></span><br><span data-ttu-id="96776-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-704">
         - TableCoercion</span></span><br><span data-ttu-id="96776-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-705">
         - TextBindings</span></span><br><span data-ttu-id="96776-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-706">
         - TextCoercion</span></span><br><span data-ttu-id="96776-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-708">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-708">Office 2016 on Mac</span></span><br><span data-ttu-id="96776-709">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-710">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="96776-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="96776-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="96776-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="96776-714">- BindingEvents</span></span><br><span data-ttu-id="96776-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-715">
         - CompressedFile</span></span><br><span data-ttu-id="96776-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="96776-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="96776-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-717">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-718">
         - File</span></span><br><span data-ttu-id="96776-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="96776-720">
         - MatrixBindings</span></span><br><span data-ttu-id="96776-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="96776-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="96776-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-723">
         - PdfFile</span></span><br><span data-ttu-id="96776-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-724">
         - Selection</span></span><br><span data-ttu-id="96776-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-725">
         - Settings</span></span><br><span data-ttu-id="96776-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="96776-726">
         - TableBindings</span></span><br><span data-ttu-id="96776-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-727">
         - TableCoercion</span></span><br><span data-ttu-id="96776-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="96776-728">
         - TextBindings</span></span><br><span data-ttu-id="96776-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-729">
         - TextCoercion</span></span><br><span data-ttu-id="96776-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="96776-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="96776-731">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="96776-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="96776-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="96776-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="96776-733">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="96776-733">Platform</span></span></th>
    <th><span data-ttu-id="96776-734">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-734">Extension points</span></span></th>
    <th><span data-ttu-id="96776-735">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="96776-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-737">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="96776-738">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-738">- Content</span></span><br><span data-ttu-id="96776-739">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-739">
         - TaskPane</span></span><br><span data-ttu-id="96776-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="96776-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="96776-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-745">- ActiveView</span></span><br><span data-ttu-id="96776-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-746">
         - CompressedFile</span></span><br><span data-ttu-id="96776-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-747">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-748">
         - File</span></span><br><span data-ttu-id="96776-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-749">
         - PdfFile</span></span><br><span data-ttu-id="96776-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-750">
         - Selection</span></span><br><span data-ttu-id="96776-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-751">
         - Settings</span></span><br><span data-ttu-id="96776-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-753">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="96776-753">Office on Windows</span></span><br><span data-ttu-id="96776-754">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-755">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-755">- Content</span></span><br><span data-ttu-id="96776-756">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-756">
         - TaskPane</span></span><br><span data-ttu-id="96776-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="96776-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="96776-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-762">- ActiveView</span></span><br><span data-ttu-id="96776-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-763">
         - CompressedFile</span></span><br><span data-ttu-id="96776-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-764">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-765">
         - File</span></span><br><span data-ttu-id="96776-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-766">
         - PdfFile</span></span><br><span data-ttu-id="96776-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-767">
         - Selection</span></span><br><span data-ttu-id="96776-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-768">
         - Settings</span></span><br><span data-ttu-id="96776-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-770">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-770">Office 2019 on Windows</span></span><br><span data-ttu-id="96776-771">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-772">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-772">- Content</span></span><br><span data-ttu-id="96776-773">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-773">
         - TaskPane</span></span><br><span data-ttu-id="96776-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-777">- ActiveView</span></span><br><span data-ttu-id="96776-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-778">
         - CompressedFile</span></span><br><span data-ttu-id="96776-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-779">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-780">
         - File</span></span><br><span data-ttu-id="96776-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-781">
         - PdfFile</span></span><br><span data-ttu-id="96776-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-782">
         - Selection</span></span><br><span data-ttu-id="96776-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-783">
         - Settings</span></span><br><span data-ttu-id="96776-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-785">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-785">Office 2016 on Windows</span></span><br><span data-ttu-id="96776-786">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-787">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-787">- Content</span></span><br><span data-ttu-id="96776-788">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="96776-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="96776-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-791">- ActiveView</span></span><br><span data-ttu-id="96776-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-792">
         - CompressedFile</span></span><br><span data-ttu-id="96776-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-793">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-794">
         - File</span></span><br><span data-ttu-id="96776-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-795">
         - PdfFile</span></span><br><span data-ttu-id="96776-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-796">
         - Selection</span></span><br><span data-ttu-id="96776-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-797">
         - Settings</span></span><br><span data-ttu-id="96776-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-799">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-799">Office 2013 on Windows</span></span><br><span data-ttu-id="96776-800">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-801">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-801">- Content</span></span><br><span data-ttu-id="96776-802">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="96776-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="96776-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="96776-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-805">- ActiveView</span></span><br><span data-ttu-id="96776-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-806">
         - CompressedFile</span></span><br><span data-ttu-id="96776-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-807">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-808">
         - File</span></span><br><span data-ttu-id="96776-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-809">
         - PdfFile</span></span><br><span data-ttu-id="96776-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-810">
         - Selection</span></span><br><span data-ttu-id="96776-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-811">
         - Settings</span></span><br><span data-ttu-id="96776-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-813">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="96776-813">Office on iPad</span></span><br><span data-ttu-id="96776-814">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-815">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-815">- Content</span></span><br><span data-ttu-id="96776-816">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="96776-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-820">- ActiveView</span></span><br><span data-ttu-id="96776-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-821">
         - CompressedFile</span></span><br><span data-ttu-id="96776-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-822">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-823">
         - File</span></span><br><span data-ttu-id="96776-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-824">
         - PdfFile</span></span><br><span data-ttu-id="96776-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-825">
         - Selection</span></span><br><span data-ttu-id="96776-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-826">
         - Settings</span></span><br><span data-ttu-id="96776-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-828">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-828">Office on Mac</span></span><br><span data-ttu-id="96776-829">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="96776-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="96776-830">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-830">- Content</span></span><br><span data-ttu-id="96776-831">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-831">
         - TaskPane</span></span><br><span data-ttu-id="96776-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="96776-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="96776-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="96776-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="96776-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-837">- ActiveView</span></span><br><span data-ttu-id="96776-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-838">
         - CompressedFile</span></span><br><span data-ttu-id="96776-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-839">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-840">
         - File</span></span><br><span data-ttu-id="96776-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-841">
         - PdfFile</span></span><br><span data-ttu-id="96776-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-842">
         - Selection</span></span><br><span data-ttu-id="96776-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-843">
         - Settings</span></span><br><span data-ttu-id="96776-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-845">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-845">Office 2019 on Mac</span></span><br><span data-ttu-id="96776-846">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-847">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-847">- Content</span></span><br><span data-ttu-id="96776-848">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-848">
         - TaskPane</span></span><br><span data-ttu-id="96776-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-852">- ActiveView</span></span><br><span data-ttu-id="96776-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-853">
         - CompressedFile</span></span><br><span data-ttu-id="96776-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-854">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-855">
         - File</span></span><br><span data-ttu-id="96776-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-856">
         - PdfFile</span></span><br><span data-ttu-id="96776-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-857">
         - Selection</span></span><br><span data-ttu-id="96776-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-858">
         - Settings</span></span><br><span data-ttu-id="96776-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-860">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="96776-860">Office 2016 on Mac</span></span><br><span data-ttu-id="96776-861">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-862">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-862">- Content</span></span><br><span data-ttu-id="96776-863">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="96776-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="96776-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="96776-866">- ActiveView</span></span><br><span data-ttu-id="96776-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="96776-867">
         - CompressedFile</span></span><br><span data-ttu-id="96776-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-868">
         - DocumentEvents</span></span><br><span data-ttu-id="96776-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="96776-869">
         - File</span></span><br><span data-ttu-id="96776-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="96776-870">
         - PdfFile</span></span><br><span data-ttu-id="96776-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="96776-871">
         - Selection</span></span><br><span data-ttu-id="96776-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-872">
         - Settings</span></span><br><span data-ttu-id="96776-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="96776-874">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="96776-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="96776-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="96776-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="96776-876">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="96776-876">Platform</span></span></th>
    <th><span data-ttu-id="96776-877">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-877">Extension points</span></span></th>
    <th><span data-ttu-id="96776-878">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="96776-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-880">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="96776-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="96776-881">- Contenu</span><span class="sxs-lookup"><span data-stu-id="96776-881">- Content</span></span><br><span data-ttu-id="96776-882">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-882">
         - TaskPane</span></span><br><span data-ttu-id="96776-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="96776-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="96776-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="96776-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="96776-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="96776-887">- DocumentEvents</span></span><br><span data-ttu-id="96776-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="96776-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="96776-889">
         - Settings</span></span><br><span data-ttu-id="96776-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="96776-891">Projet</span><span class="sxs-lookup"><span data-stu-id="96776-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="96776-892">Plateforme</span><span class="sxs-lookup"><span data-stu-id="96776-892">Platform</span></span></th>
    <th><span data-ttu-id="96776-893">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="96776-893">Extension points</span></span></th>
    <th><span data-ttu-id="96776-894">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="96776-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="96776-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-896">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-896">Office 2019 on Windows</span></span><br><span data-ttu-id="96776-897">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-898">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="96776-900">- Selection</span></span><br><span data-ttu-id="96776-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-902">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-902">Office 2016 on Windows</span></span><br><span data-ttu-id="96776-903">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-904">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="96776-906">- Selection</span></span><br><span data-ttu-id="96776-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="96776-908">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="96776-908">Office 2013 on Windows</span></span><br><span data-ttu-id="96776-909">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="96776-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="96776-910">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="96776-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="96776-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="96776-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="96776-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="96776-912">- Selection</span></span><br><span data-ttu-id="96776-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="96776-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="96776-914">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="96776-914">See also</span></span>

- [<span data-ttu-id="96776-915">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="96776-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="96776-916">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="96776-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="96776-917">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="96776-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="96776-918">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="96776-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="96776-919">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="96776-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="96776-920">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="96776-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="96776-921">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="96776-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="96776-922">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="96776-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="96776-923">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="96776-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="96776-924">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="96776-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="96776-925">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="96776-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="96776-926">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="96776-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)