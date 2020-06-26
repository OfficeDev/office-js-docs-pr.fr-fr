---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 06/23/2020
localization_priority: Priority
ms.openlocfilehash: 979c873b1c5f2d1d7847414f037d5c75737aa33d
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888158"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ef9d3-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="ef9d3-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ef9d3-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span><span class="sxs-lookup"><span data-stu-id="ef9d3-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="ef9d3-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span><span class="sxs-lookup"><span data-stu-id="ef9d3-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ef9d3-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="ef9d3-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ef9d3-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="ef9d3-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ef9d3-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ef9d3-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ef9d3-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ef9d3-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ef9d3-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ef9d3-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ef9d3-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-114">- TaskPane</span></span><br><span data-ttu-id="ef9d3-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-115">
        - Content</span></span><br><span data-ttu-id="ef9d3-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-116">
        - Custom Functions</span></span><br><span data-ttu-id="ef9d3-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="ef9d3-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ef9d3-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ef9d3-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ef9d3-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ef9d3-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ef9d3-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-131">
        - BindingEvents</span></span><br><span data-ttu-id="ef9d3-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-132">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-133">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-134">
        - File</span></span><br><span data-ttu-id="ef9d3-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-135">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-137">
        - Selection</span></span><br><span data-ttu-id="ef9d3-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-138">
        - Settings</span></span><br><span data-ttu-id="ef9d3-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-139">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-140">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-141">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-143">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-143">Office on Windows</span></span><br><span data-ttu-id="ef9d3-144">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-145">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-145">- TaskPane</span></span><br><span data-ttu-id="ef9d3-146">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-146">
        - Content</span></span><br><span data-ttu-id="ef9d3-147">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-147">
        - Custom Functions</span></span><br><span data-ttu-id="ef9d3-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="ef9d3-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ef9d3-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ef9d3-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ef9d3-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ef9d3-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ef9d3-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-163">
        - BindingEvents</span></span><br><span data-ttu-id="ef9d3-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-164">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-165">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-166">
        - File</span></span><br><span data-ttu-id="ef9d3-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-167">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-169">
        - Selection</span></span><br><span data-ttu-id="ef9d3-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-170">
        - Settings</span></span><br><span data-ttu-id="ef9d3-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-171">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-172">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-173">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-175">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-175">Office 2019 on Windows</span></span><br><span data-ttu-id="ef9d3-176">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ef9d3-177">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-177">- TaskPane</span></span><br><span data-ttu-id="ef9d3-178">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-178">
        - Content</span></span><br><span data-ttu-id="ef9d3-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ef9d3-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-190">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-191">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-192">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-193">
        - File</span></span><br><span data-ttu-id="ef9d3-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-194">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-196">
        - Selection</span></span><br><span data-ttu-id="ef9d3-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-197">
        - Settings</span></span><br><span data-ttu-id="ef9d3-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-198">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-199">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-200">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-202">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-202">Office 2016 on Windows</span></span><br><span data-ttu-id="ef9d3-203">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ef9d3-204">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-204">- TaskPane</span></span><br><span data-ttu-id="ef9d3-205">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-205">
        - Content</span></span></td>
    <td><span data-ttu-id="ef9d3-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ef9d3-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-209">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-210">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-211">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-212">
        - File</span></span><br><span data-ttu-id="ef9d3-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-213">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-215">
        - Selection</span></span><br><span data-ttu-id="ef9d3-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-216">
        - Settings</span></span><br><span data-ttu-id="ef9d3-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-217">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-218">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-219">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-221">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-221">Office 2013 on Windows</span></span><br><span data-ttu-id="ef9d3-222">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ef9d3-223">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-223">
        - TaskPane</span></span><br><span data-ttu-id="ef9d3-224">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ef9d3-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ef9d3-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-227">
        - BindingEvents</span></span><br><span data-ttu-id="ef9d3-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-228">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-229">
        - File</span></span><br><span data-ttu-id="ef9d3-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-230">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-232">
        - Selection</span></span><br><span data-ttu-id="ef9d3-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-233">
        - Settings</span></span><br><span data-ttu-id="ef9d3-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-234">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-235">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-236">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-238">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="ef9d3-238">Office on iPad</span></span><br><span data-ttu-id="ef9d3-239">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ef9d3-240">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-240">- TaskPane</span></span><br><span data-ttu-id="ef9d3-241">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-241">
        - Content</span></span></td>
    <td><span data-ttu-id="ef9d3-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ef9d3-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ef9d3-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ef9d3-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-255">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-256">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-257">
        - File</span></span><br><span data-ttu-id="ef9d3-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-258">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-260">
        - Selection</span></span><br><span data-ttu-id="ef9d3-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-261">
        - Settings</span></span><br><span data-ttu-id="ef9d3-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-262">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-263">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-264">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-266">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-266">Office on Mac</span></span><br><span data-ttu-id="ef9d3-267">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ef9d3-268">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-268">- TaskPane</span></span><br><span data-ttu-id="ef9d3-269">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-269">
        - Content</span></span><br><span data-ttu-id="ef9d3-270">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-270">
        - Custom Functions</span></span><br><span data-ttu-id="ef9d3-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ef9d3-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ef9d3-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ef9d3-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ef9d3-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ef9d3-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-286">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-287">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-288">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-289">
        - File</span></span><br><span data-ttu-id="ef9d3-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-290">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-292">
        - PdfFile</span></span><br><span data-ttu-id="ef9d3-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-293">
        - Selection</span></span><br><span data-ttu-id="ef9d3-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-294">
        - Settings</span></span><br><span data-ttu-id="ef9d3-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-295">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-296">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-297">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-299">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-299">Office 2019 on Mac</span></span><br><span data-ttu-id="ef9d3-300">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ef9d3-301">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-301">- TaskPane</span></span><br><span data-ttu-id="ef9d3-302">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-302">
        - Content</span></span><br><span data-ttu-id="ef9d3-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ef9d3-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ef9d3-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ef9d3-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ef9d3-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ef9d3-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ef9d3-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-314">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-315">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-316">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-317">
        - File</span></span><br><span data-ttu-id="ef9d3-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-318">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-320">
        - PdfFile</span></span><br><span data-ttu-id="ef9d3-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-321">
        - Selection</span></span><br><span data-ttu-id="ef9d3-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-322">
        - Settings</span></span><br><span data-ttu-id="ef9d3-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-323">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-324">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-325">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-327">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-327">Office 2016 on Mac</span></span><br><span data-ttu-id="ef9d3-328">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ef9d3-329">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-329">- TaskPane</span></span><br><span data-ttu-id="ef9d3-330">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-330">
        - Content</span></span></td>
    <td><span data-ttu-id="ef9d3-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ef9d3-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ef9d3-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-334">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-335">
        - CompressedFile</span></span><br><span data-ttu-id="ef9d3-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-336">
        - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-337">
        - File</span></span><br><span data-ttu-id="ef9d3-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-338">
        - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-340">
        - PdfFile</span></span><br><span data-ttu-id="ef9d3-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-341">
        - Selection</span></span><br><span data-ttu-id="ef9d3-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-342">
        - Settings</span></span><br><span data-ttu-id="ef9d3-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-343">
        - TableBindings</span></span><br><span data-ttu-id="ef9d3-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-344">
        - TableCoercion</span></span><br><span data-ttu-id="ef9d3-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-345">
        - TextBindings</span></span><br><span data-ttu-id="ef9d3-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ef9d3-347">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ef9d3-348">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ef9d3-349">Plateforme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ef9d3-350">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ef9d3-351">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ef9d3-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-353">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-353">Office on the web</span></span></td>
    <td><span data-ttu-id="ef9d3-354">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ef9d3-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-356">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-356">Office on Windows</span></span><br><span data-ttu-id="ef9d3-357">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ef9d3-358">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ef9d3-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-360">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-360">Office on Mac</span></span><br><span data-ttu-id="ef9d3-361">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ef9d3-362">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ef9d3-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ef9d3-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ef9d3-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="ef9d3-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ef9d3-365">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-365">Platform</span></span></th>
    <th><span data-ttu-id="ef9d3-366">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-366">Extension points</span></span></th>
    <th><span data-ttu-id="ef9d3-367">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="ef9d3-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-369">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-369">Office on the web</span></span><br><span data-ttu-id="ef9d3-370">(moderne)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-370">(modern)</span></span></td>
    <td> <span data-ttu-id="ef9d3-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ef9d3-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ef9d3-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ef9d3-384">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-385">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-385">Office on the web</span></span><br><span data-ttu-id="ef9d3-386">(classique)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-386">(classic)</span></span></td>
    <td> <span data-ttu-id="ef9d3-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ef9d3-398">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-399">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-399">Office on Windows</span></span><br><span data-ttu-id="ef9d3-400">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ef9d3-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ef9d3-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ef9d3-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ef9d3-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-416">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-416">Office 2019 on Windows</span></span><br><span data-ttu-id="ef9d3-417">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ef9d3-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ef9d3-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ef9d3-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-432">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-432">Office 2016 on Windows</span></span><br><span data-ttu-id="ef9d3-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ef9d3-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ef9d3-444">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-445">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-445">Office 2013 on Windows</span></span><br><span data-ttu-id="ef9d3-446">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="ef9d3-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ef9d3-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ef9d3-455">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-456">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="ef9d3-456">Office on iOS</span></span><br><span data-ttu-id="ef9d3-457">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ef9d3-465">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-466">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-466">Office on Mac</span></span><br><span data-ttu-id="ef9d3-467">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ef9d3-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ef9d3-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ef9d3-481">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-482">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-482">Office 2019 on Mac</span></span><br><span data-ttu-id="ef9d3-483">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ef9d3-495">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-496">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-496">Office 2016 on Mac</span></span><br><span data-ttu-id="ef9d3-497">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ef9d3-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ef9d3-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ef9d3-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ef9d3-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ef9d3-509">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-510">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="ef9d3-510">Office on Android</span></span><br><span data-ttu-id="ef9d3-511">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ef9d3-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="ef9d3-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ef9d3-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ef9d3-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ef9d3-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ef9d3-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ef9d3-520">Non disponible</span><span class="sxs-lookup"><span data-stu-id="ef9d3-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ef9d3-521">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef9d3-522">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="ef9d3-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ef9d3-523">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="ef9d3-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ef9d3-524">Word</span><span class="sxs-lookup"><span data-stu-id="ef9d3-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ef9d3-525">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-525">Platform</span></span></th>
    <th><span data-ttu-id="ef9d3-526">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-526">Extension points</span></span></th>
    <th><span data-ttu-id="ef9d3-527">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="ef9d3-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-529">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="ef9d3-530">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-530">- TaskPane</span></span><br><span data-ttu-id="ef9d3-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-538">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-540">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-541">
         - File</span></span><br><span data-ttu-id="ef9d3-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-543">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-546">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-547">
         - Selection</span></span><br><span data-ttu-id="ef9d3-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-548">
         - Settings</span></span><br><span data-ttu-id="ef9d3-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-549">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-550">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-551">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-552">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-554">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-554">Office on Windows</span></span><br><span data-ttu-id="ef9d3-555">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-556">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-556">- TaskPane</span></span><br><span data-ttu-id="ef9d3-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-564">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-565">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-567">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-568">
         - File</span></span><br><span data-ttu-id="ef9d3-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-570">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-573">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-574">
         - Selection</span></span><br><span data-ttu-id="ef9d3-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-575">
         - Settings</span></span><br><span data-ttu-id="ef9d3-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-576">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-577">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-578">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-579">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-581">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-581">Office 2019 on Windows</span></span><br><span data-ttu-id="ef9d3-582">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-583">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="ef9d3-583">- TaskPane</span></span><br><span data-ttu-id="ef9d3-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-590">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-591">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-593">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-594">
         - File</span></span><br><span data-ttu-id="ef9d3-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-596">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-599">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-600">
         - Selection</span></span><br><span data-ttu-id="ef9d3-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-601">
         - Settings</span></span><br><span data-ttu-id="ef9d3-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-602">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-603">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-604">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-605">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-607">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-607">Office 2016 on Windows</span></span><br><span data-ttu-id="ef9d3-608">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-609">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ef9d3-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-613">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-614">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-616">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-617">
         - File</span></span><br><span data-ttu-id="ef9d3-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-619">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-622">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-623">
         - Selection</span></span><br><span data-ttu-id="ef9d3-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-624">
         - Settings</span></span><br><span data-ttu-id="ef9d3-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-625">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-626">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-627">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-628">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-630">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-630">Office 2013 on Windows</span></span><br><span data-ttu-id="ef9d3-631">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ef9d3-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-635">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-636">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-638">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-639">
         - File</span></span><br><span data-ttu-id="ef9d3-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-641">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-644">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-645">
         - Selection</span></span><br><span data-ttu-id="ef9d3-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-646">
         - Settings</span></span><br><span data-ttu-id="ef9d3-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-647">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-648">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-649">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-650">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-652">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="ef9d3-652">Office on iPad</span></span><br><span data-ttu-id="ef9d3-653">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-654">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ef9d3-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-660">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-661">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-663">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-664">
         - File</span></span><br><span data-ttu-id="ef9d3-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-666">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-669">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-670">
         - Selection</span></span><br><span data-ttu-id="ef9d3-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-671">
         - Settings</span></span><br><span data-ttu-id="ef9d3-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-672">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-673">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-674">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-675">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-677">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-677">Office on Mac</span></span><br><span data-ttu-id="ef9d3-678">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-679">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-679">- TaskPane</span></span><br><span data-ttu-id="ef9d3-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ef9d3-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-687">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-688">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-690">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-691">
         - File</span></span><br><span data-ttu-id="ef9d3-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-693">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-696">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-697">
         - Selection</span></span><br><span data-ttu-id="ef9d3-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-698">
         - Settings</span></span><br><span data-ttu-id="ef9d3-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-699">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-700">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-701">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-702">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-704">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-704">Office 2019 on Mac</span></span><br><span data-ttu-id="ef9d3-705">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-706">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="ef9d3-706">- TaskPane</span></span><br><span data-ttu-id="ef9d3-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ef9d3-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ef9d3-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ef9d3-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-713">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-714">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-716">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-717">
         - File</span></span><br><span data-ttu-id="ef9d3-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-719">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-722">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-723">
         - Selection</span></span><br><span data-ttu-id="ef9d3-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-724">
         - Settings</span></span><br><span data-ttu-id="ef9d3-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-725">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-726">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-727">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-728">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-730">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-730">Office 2016 on Mac</span></span><br><span data-ttu-id="ef9d3-731">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-732">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ef9d3-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-736">- BindingEvents</span></span><br><span data-ttu-id="ef9d3-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-737">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ef9d3-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="ef9d3-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-739">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-740">
         - File</span></span><br><span data-ttu-id="ef9d3-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-742">
         - MatrixBindings</span></span><br><span data-ttu-id="ef9d3-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="ef9d3-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ef9d3-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-745">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-746">
         - Selection</span></span><br><span data-ttu-id="ef9d3-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-747">
         - Settings</span></span><br><span data-ttu-id="ef9d3-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-748">
         - TableBindings</span></span><br><span data-ttu-id="ef9d3-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-749">
         - TableCoercion</span></span><br><span data-ttu-id="ef9d3-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-750">
         - TextBindings</span></span><br><span data-ttu-id="ef9d3-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-751">
         - TextCoercion</span></span><br><span data-ttu-id="ef9d3-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ef9d3-753">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ef9d3-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ef9d3-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ef9d3-755">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-755">Platform</span></span></th>
    <th><span data-ttu-id="ef9d3-756">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-756">Extension points</span></span></th>
    <th><span data-ttu-id="ef9d3-757">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="ef9d3-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-759">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="ef9d3-760">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-760">- Content</span></span><br><span data-ttu-id="ef9d3-761">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-761">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-767">- ActiveView</span></span><br><span data-ttu-id="ef9d3-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-768">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-769">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-770">
         - File</span></span><br><span data-ttu-id="ef9d3-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-771">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-772">
         - Selection</span></span><br><span data-ttu-id="ef9d3-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-773">
         - Settings</span></span><br><span data-ttu-id="ef9d3-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-775">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-775">Office on Windows</span></span><br><span data-ttu-id="ef9d3-776">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-777">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-777">- Content</span></span><br><span data-ttu-id="ef9d3-778">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-778">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-784">- ActiveView</span></span><br><span data-ttu-id="ef9d3-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-785">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-786">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-787">
         - File</span></span><br><span data-ttu-id="ef9d3-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-788">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-789">
         - Selection</span></span><br><span data-ttu-id="ef9d3-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-790">
         - Settings</span></span><br><span data-ttu-id="ef9d3-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-792">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-792">Office 2019 on Windows</span></span><br><span data-ttu-id="ef9d3-793">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-794">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-794">- Content</span></span><br><span data-ttu-id="ef9d3-795">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-795">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-799">- ActiveView</span></span><br><span data-ttu-id="ef9d3-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-800">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-801">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-802">
         - File</span></span><br><span data-ttu-id="ef9d3-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-803">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-804">
         - Selection</span></span><br><span data-ttu-id="ef9d3-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-805">
         - Settings</span></span><br><span data-ttu-id="ef9d3-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-807">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-807">Office 2016 on Windows</span></span><br><span data-ttu-id="ef9d3-808">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-809">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-809">- Content</span></span><br><span data-ttu-id="ef9d3-810">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ef9d3-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-813">- ActiveView</span></span><br><span data-ttu-id="ef9d3-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-814">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-815">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-816">
         - File</span></span><br><span data-ttu-id="ef9d3-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-817">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-818">
         - Selection</span></span><br><span data-ttu-id="ef9d3-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-819">
         - Settings</span></span><br><span data-ttu-id="ef9d3-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-821">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-821">Office 2013 on Windows</span></span><br><span data-ttu-id="ef9d3-822">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-823">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-823">- Content</span></span><br><span data-ttu-id="ef9d3-824">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ef9d3-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ef9d3-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-827">- ActiveView</span></span><br><span data-ttu-id="ef9d3-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-828">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-829">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-830">
         - File</span></span><br><span data-ttu-id="ef9d3-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-831">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-832">
         - Selection</span></span><br><span data-ttu-id="ef9d3-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-833">
         - Settings</span></span><br><span data-ttu-id="ef9d3-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-835">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="ef9d3-835">Office on iPad</span></span><br><span data-ttu-id="ef9d3-836">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-837">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-837">- Content</span></span><br><span data-ttu-id="ef9d3-838">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-842">- ActiveView</span></span><br><span data-ttu-id="ef9d3-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-843">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-844">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-845">
         - File</span></span><br><span data-ttu-id="ef9d3-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-846">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-847">
         - Selection</span></span><br><span data-ttu-id="ef9d3-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-848">
         - Settings</span></span><br><span data-ttu-id="ef9d3-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-850">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-850">Office on Mac</span></span><br><span data-ttu-id="ef9d3-851">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ef9d3-852">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-852">- Content</span></span><br><span data-ttu-id="ef9d3-853">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-853">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ef9d3-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-859">- ActiveView</span></span><br><span data-ttu-id="ef9d3-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-860">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-861">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-862">
         - File</span></span><br><span data-ttu-id="ef9d3-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-863">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-864">
         - Selection</span></span><br><span data-ttu-id="ef9d3-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-865">
         - Settings</span></span><br><span data-ttu-id="ef9d3-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-867">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-867">Office 2019 on Mac</span></span><br><span data-ttu-id="ef9d3-868">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-869">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-869">- Content</span></span><br><span data-ttu-id="ef9d3-870">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-870">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-874">- ActiveView</span></span><br><span data-ttu-id="ef9d3-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-875">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-876">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-877">
         - File</span></span><br><span data-ttu-id="ef9d3-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-878">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-879">
         - Selection</span></span><br><span data-ttu-id="ef9d3-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-880">
         - Settings</span></span><br><span data-ttu-id="ef9d3-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-882">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-882">Office 2016 on Mac</span></span><br><span data-ttu-id="ef9d3-883">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-884">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-884">- Content</span></span><br><span data-ttu-id="ef9d3-885">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ef9d3-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ef9d3-888">- ActiveView</span></span><br><span data-ttu-id="ef9d3-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-889">
         - CompressedFile</span></span><br><span data-ttu-id="ef9d3-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-890">
         - DocumentEvents</span></span><br><span data-ttu-id="ef9d3-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="ef9d3-891">
         - File</span></span><br><span data-ttu-id="ef9d3-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ef9d3-892">
         - PdfFile</span></span><br><span data-ttu-id="ef9d3-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-893">
         - Selection</span></span><br><span data-ttu-id="ef9d3-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-894">
         - Settings</span></span><br><span data-ttu-id="ef9d3-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ef9d3-896">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="ef9d3-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ef9d3-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="ef9d3-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ef9d3-898">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-898">Platform</span></span></th>
    <th><span data-ttu-id="ef9d3-899">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-899">Extension points</span></span></th>
    <th><span data-ttu-id="ef9d3-900">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="ef9d3-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-902">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="ef9d3-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="ef9d3-903">- Contenu</span><span class="sxs-lookup"><span data-stu-id="ef9d3-903">- Content</span></span><br><span data-ttu-id="ef9d3-904">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-904">
         - TaskPane</span></span><br><span data-ttu-id="ef9d3-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ef9d3-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ef9d3-909">- DocumentEvents</span></span><br><span data-ttu-id="ef9d3-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="ef9d3-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ef9d3-911">
         - Settings</span></span><br><span data-ttu-id="ef9d3-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ef9d3-913">Projet</span><span class="sxs-lookup"><span data-stu-id="ef9d3-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ef9d3-914">Plateforme</span><span class="sxs-lookup"><span data-stu-id="ef9d3-914">Platform</span></span></th>
    <th><span data-ttu-id="ef9d3-915">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="ef9d3-915">Extension points</span></span></th>
    <th><span data-ttu-id="ef9d3-916">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="ef9d3-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-918">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-918">Office 2019 on Windows</span></span><br><span data-ttu-id="ef9d3-919">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-920">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-922">- Selection</span></span><br><span data-ttu-id="ef9d3-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-924">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-924">Office 2016 on Windows</span></span><br><span data-ttu-id="ef9d3-925">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-926">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-928">- Selection</span></span><br><span data-ttu-id="ef9d3-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ef9d3-930">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="ef9d3-930">Office 2013 on Windows</span></span><br><span data-ttu-id="ef9d3-931">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ef9d3-932">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ef9d3-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ef9d3-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ef9d3-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="ef9d3-934">- Selection</span></span><br><span data-ttu-id="ef9d3-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ef9d3-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ef9d3-936">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ef9d3-936">See also</span></span>

- [<span data-ttu-id="ef9d3-937">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ef9d3-938">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ef9d3-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ef9d3-939">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="ef9d3-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ef9d3-940">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="ef9d3-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ef9d3-941">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="ef9d3-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ef9d3-942">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="ef9d3-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ef9d3-943">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ef9d3-944">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ef9d3-945">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ef9d3-946">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ef9d3-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ef9d3-947">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="ef9d3-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ef9d3-948">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="ef9d3-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)