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
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d1a41-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="d1a41-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d1a41-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="d1a41-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d1a41-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="d1a41-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="d1a41-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="d1a41-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="d1a41-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d1a41-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d1a41-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d1a41-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d1a41-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d1a41-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d1a41-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="d1a41-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-114">- TaskPane</span></span><br><span data-ttu-id="d1a41-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-115">
        - Content</span></span><br><span data-ttu-id="d1a41-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-116">
        - Custom Functions</span></span><br><span data-ttu-id="d1a41-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="d1a41-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d1a41-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d1a41-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d1a41-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d1a41-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="d1a41-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-131">
        - BindingEvents</span></span><br><span data-ttu-id="d1a41-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-132">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-133">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-134">
        - File</span></span><br><span data-ttu-id="d1a41-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-135">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-137">
        - Selection</span></span><br><span data-ttu-id="d1a41-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-138">
        - Settings</span></span><br><span data-ttu-id="d1a41-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-139">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-140">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-141">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-143">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-143">Office on Windows</span></span><br><span data-ttu-id="d1a41-144">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-145">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-145">- TaskPane</span></span><br><span data-ttu-id="d1a41-146">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-146">
        - Content</span></span><br><span data-ttu-id="d1a41-147">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-147">
        - Custom Functions</span></span><br><span data-ttu-id="d1a41-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="d1a41-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d1a41-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d1a41-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d1a41-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d1a41-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d1a41-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-163">
        - BindingEvents</span></span><br><span data-ttu-id="d1a41-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-164">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-165">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-166">
        - File</span></span><br><span data-ttu-id="d1a41-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-167">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-169">
        - Selection</span></span><br><span data-ttu-id="d1a41-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-170">
        - Settings</span></span><br><span data-ttu-id="d1a41-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-171">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-172">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-173">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-175">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-175">Office 2019 on Windows</span></span><br><span data-ttu-id="d1a41-176">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d1a41-177">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-177">- TaskPane</span></span><br><span data-ttu-id="d1a41-178">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-178">
        - Content</span></span><br><span data-ttu-id="d1a41-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d1a41-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-190">- BindingEvents</span></span><br><span data-ttu-id="d1a41-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-191">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-192">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-193">
        - File</span></span><br><span data-ttu-id="d1a41-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-194">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-196">
        - Selection</span></span><br><span data-ttu-id="d1a41-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-197">
        - Settings</span></span><br><span data-ttu-id="d1a41-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-198">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-199">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-200">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-202">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-202">Office 2016 on Windows</span></span><br><span data-ttu-id="d1a41-203">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d1a41-204">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-204">- TaskPane</span></span><br><span data-ttu-id="d1a41-205">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-205">
        - Content</span></span></td>
    <td><span data-ttu-id="d1a41-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d1a41-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-209">- BindingEvents</span></span><br><span data-ttu-id="d1a41-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-210">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-211">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-212">
        - File</span></span><br><span data-ttu-id="d1a41-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-213">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-215">
        - Selection</span></span><br><span data-ttu-id="d1a41-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-216">
        - Settings</span></span><br><span data-ttu-id="d1a41-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-217">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-218">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-219">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-221">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-221">Office 2013 on Windows</span></span><br><span data-ttu-id="d1a41-222">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d1a41-223">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-223">
        - TaskPane</span></span><br><span data-ttu-id="d1a41-224">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d1a41-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d1a41-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d1a41-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-227">
        - BindingEvents</span></span><br><span data-ttu-id="d1a41-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-228">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-229">
        - File</span></span><br><span data-ttu-id="d1a41-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-230">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-232">
        - Selection</span></span><br><span data-ttu-id="d1a41-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-233">
        - Settings</span></span><br><span data-ttu-id="d1a41-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-234">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-235">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-236">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-238">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d1a41-238">Office on iPad</span></span><br><span data-ttu-id="d1a41-239">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d1a41-240">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-240">- TaskPane</span></span><br><span data-ttu-id="d1a41-241">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-241">
        - Content</span></span></td>
    <td><span data-ttu-id="d1a41-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d1a41-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d1a41-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d1a41-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-255">- BindingEvents</span></span><br><span data-ttu-id="d1a41-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-256">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-257">
        - File</span></span><br><span data-ttu-id="d1a41-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-258">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-260">
        - Selection</span></span><br><span data-ttu-id="d1a41-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-261">
        - Settings</span></span><br><span data-ttu-id="d1a41-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-262">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-263">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-264">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-266">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-266">Office on Mac</span></span><br><span data-ttu-id="d1a41-267">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d1a41-268">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-268">- TaskPane</span></span><br><span data-ttu-id="d1a41-269">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-269">
        - Content</span></span><br><span data-ttu-id="d1a41-270">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-270">
        - Custom Functions</span></span><br><span data-ttu-id="d1a41-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d1a41-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d1a41-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d1a41-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d1a41-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d1a41-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-286">- BindingEvents</span></span><br><span data-ttu-id="d1a41-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-287">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-288">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-289">
        - File</span></span><br><span data-ttu-id="d1a41-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-290">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-292">
        - PdfFile</span></span><br><span data-ttu-id="d1a41-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-293">
        - Selection</span></span><br><span data-ttu-id="d1a41-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-294">
        - Settings</span></span><br><span data-ttu-id="d1a41-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-295">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-296">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-297">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-299">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-299">Office 2019 on Mac</span></span><br><span data-ttu-id="d1a41-300">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d1a41-301">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-301">- TaskPane</span></span><br><span data-ttu-id="d1a41-302">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-302">
        - Content</span></span><br><span data-ttu-id="d1a41-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d1a41-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d1a41-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d1a41-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d1a41-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d1a41-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d1a41-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d1a41-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d1a41-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-314">- BindingEvents</span></span><br><span data-ttu-id="d1a41-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-315">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-316">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-317">
        - File</span></span><br><span data-ttu-id="d1a41-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-318">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-320">
        - PdfFile</span></span><br><span data-ttu-id="d1a41-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-321">
        - Selection</span></span><br><span data-ttu-id="d1a41-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-322">
        - Settings</span></span><br><span data-ttu-id="d1a41-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-323">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-324">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-325">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-327">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-327">Office 2016 on Mac</span></span><br><span data-ttu-id="d1a41-328">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d1a41-329">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-329">- TaskPane</span></span><br><span data-ttu-id="d1a41-330">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-330">
        - Content</span></span></td>
    <td><span data-ttu-id="d1a41-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d1a41-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d1a41-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d1a41-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-334">- BindingEvents</span></span><br><span data-ttu-id="d1a41-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-335">
        - CompressedFile</span></span><br><span data-ttu-id="d1a41-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-336">
        - DocumentEvents</span></span><br><span data-ttu-id="d1a41-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-337">
        - File</span></span><br><span data-ttu-id="d1a41-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-338">
        - MatrixBindings</span></span><br><span data-ttu-id="d1a41-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-340">
        - PdfFile</span></span><br><span data-ttu-id="d1a41-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-341">
        - Selection</span></span><br><span data-ttu-id="d1a41-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-342">
        - Settings</span></span><br><span data-ttu-id="d1a41-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-343">
        - TableBindings</span></span><br><span data-ttu-id="d1a41-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-344">
        - TableCoercion</span></span><br><span data-ttu-id="d1a41-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-345">
        - TextBindings</span></span><br><span data-ttu-id="d1a41-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d1a41-347">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d1a41-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="d1a41-348">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="d1a41-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d1a41-349">Plateforme</span><span class="sxs-lookup"><span data-stu-id="d1a41-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d1a41-350">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d1a41-351">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d1a41-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-353">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-353">Office on the web</span></span></td>
    <td><span data-ttu-id="d1a41-354">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d1a41-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-356">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-356">Office on Windows</span></span><br><span data-ttu-id="d1a41-357">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d1a41-358">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d1a41-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-360">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-360">Office on Mac</span></span><br><span data-ttu-id="d1a41-361">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="d1a41-362">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d1a41-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d1a41-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="d1a41-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="d1a41-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d1a41-365">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d1a41-365">Platform</span></span></th>
    <th><span data-ttu-id="d1a41-366">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-366">Extension points</span></span></th>
    <th><span data-ttu-id="d1a41-367">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="d1a41-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-369">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-369">Office on the web</span></span><br><span data-ttu-id="d1a41-370">(moderne)</span><span class="sxs-lookup"><span data-stu-id="d1a41-370">(modern)</span></span></td>
    <td> <span data-ttu-id="d1a41-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d1a41-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d1a41-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d1a41-384">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-385">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-385">Office on the web</span></span><br><span data-ttu-id="d1a41-386">(classique)</span><span class="sxs-lookup"><span data-stu-id="d1a41-386">(classic)</span></span></td>
    <td> <span data-ttu-id="d1a41-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d1a41-398">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-399">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-399">Office on Windows</span></span><br><span data-ttu-id="d1a41-400">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d1a41-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d1a41-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d1a41-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d1a41-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d1a41-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-416">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-416">Office 2019 on Windows</span></span><br><span data-ttu-id="d1a41-417">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d1a41-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d1a41-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d1a41-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d1a41-431">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-432">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-432">Office 2016 on Windows</span></span><br><span data-ttu-id="d1a41-433">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d1a41-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d1a41-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d1a41-444">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-445">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-445">Office 2013 on Windows</span></span><br><span data-ttu-id="d1a41-446">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="d1a41-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="d1a41-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d1a41-455">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-456">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="d1a41-456">Office on iOS</span></span><br><span data-ttu-id="d1a41-457">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d1a41-465">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-466">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-466">Office on Mac</span></span><br><span data-ttu-id="d1a41-467">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d1a41-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d1a41-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d1a41-481">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-482">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-482">Office 2019 on Mac</span></span><br><span data-ttu-id="d1a41-483">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d1a41-495">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-496">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-496">Office 2016 on Mac</span></span><br><span data-ttu-id="d1a41-497">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d1a41-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d1a41-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d1a41-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d1a41-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d1a41-509">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-510">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="d1a41-510">Office on Android</span></span><br><span data-ttu-id="d1a41-511">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d1a41-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)</span><span class="sxs-lookup"><span data-stu-id="d1a41-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="d1a41-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d1a41-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d1a41-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d1a41-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d1a41-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d1a41-520">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d1a41-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="d1a41-521">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d1a41-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d1a41-522">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d1a41-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="d1a41-523">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="d1a41-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="d1a41-524">Word</span><span class="sxs-lookup"><span data-stu-id="d1a41-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d1a41-525">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d1a41-525">Platform</span></span></th>
    <th><span data-ttu-id="d1a41-526">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-526">Extension points</span></span></th>
    <th><span data-ttu-id="d1a41-527">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="d1a41-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-529">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="d1a41-530">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-530">- TaskPane</span></span><br><span data-ttu-id="d1a41-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d1a41-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-538">- BindingEvents</span></span><br><span data-ttu-id="d1a41-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-540">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-541">
         - File</span></span><br><span data-ttu-id="d1a41-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-543">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-546">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-547">
         - Selection</span></span><br><span data-ttu-id="d1a41-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-548">
         - Settings</span></span><br><span data-ttu-id="d1a41-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-549">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-550">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-551">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-552">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-554">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-554">Office on Windows</span></span><br><span data-ttu-id="d1a41-555">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-556">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-556">- TaskPane</span></span><br><span data-ttu-id="d1a41-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d1a41-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-564">- BindingEvents</span></span><br><span data-ttu-id="d1a41-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-565">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-567">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-568">
         - File</span></span><br><span data-ttu-id="d1a41-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-570">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-573">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-574">
         - Selection</span></span><br><span data-ttu-id="d1a41-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-575">
         - Settings</span></span><br><span data-ttu-id="d1a41-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-576">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-577">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-578">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-579">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-581">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-581">Office 2019 on Windows</span></span><br><span data-ttu-id="d1a41-582">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-583">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="d1a41-583">- TaskPane</span></span><br><span data-ttu-id="d1a41-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-590">- BindingEvents</span></span><br><span data-ttu-id="d1a41-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-591">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-593">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-594">
         - File</span></span><br><span data-ttu-id="d1a41-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-596">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-599">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-600">
         - Selection</span></span><br><span data-ttu-id="d1a41-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-601">
         - Settings</span></span><br><span data-ttu-id="d1a41-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-602">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-603">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-604">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-605">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-607">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-607">Office 2016 on Windows</span></span><br><span data-ttu-id="d1a41-608">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-609">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d1a41-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-613">- BindingEvents</span></span><br><span data-ttu-id="d1a41-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-614">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-616">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-617">
         - File</span></span><br><span data-ttu-id="d1a41-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-619">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-622">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-623">
         - Selection</span></span><br><span data-ttu-id="d1a41-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-624">
         - Settings</span></span><br><span data-ttu-id="d1a41-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-625">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-626">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-627">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-628">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-630">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-630">Office 2013 on Windows</span></span><br><span data-ttu-id="d1a41-631">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-632">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d1a41-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d1a41-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-635">- BindingEvents</span></span><br><span data-ttu-id="d1a41-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-636">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-638">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-639">
         - File</span></span><br><span data-ttu-id="d1a41-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-641">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-644">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-645">
         - Selection</span></span><br><span data-ttu-id="d1a41-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-646">
         - Settings</span></span><br><span data-ttu-id="d1a41-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-647">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-648">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-649">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-650">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-652">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d1a41-652">Office on iPad</span></span><br><span data-ttu-id="d1a41-653">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-654">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d1a41-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-660">- BindingEvents</span></span><br><span data-ttu-id="d1a41-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-661">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-663">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-664">
         - File</span></span><br><span data-ttu-id="d1a41-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-666">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-669">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-670">
         - Selection</span></span><br><span data-ttu-id="d1a41-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-671">
         - Settings</span></span><br><span data-ttu-id="d1a41-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-672">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-673">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-674">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-675">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-677">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-677">Office on Mac</span></span><br><span data-ttu-id="d1a41-678">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-679">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-679">- TaskPane</span></span><br><span data-ttu-id="d1a41-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="d1a41-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-687">- BindingEvents</span></span><br><span data-ttu-id="d1a41-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-688">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-690">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-691">
         - File</span></span><br><span data-ttu-id="d1a41-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-693">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-696">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-697">
         - Selection</span></span><br><span data-ttu-id="d1a41-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-698">
         - Settings</span></span><br><span data-ttu-id="d1a41-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-699">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-700">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-701">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-702">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-704">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-704">Office 2019 on Mac</span></span><br><span data-ttu-id="d1a41-705">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-706">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="d1a41-706">- TaskPane</span></span><br><span data-ttu-id="d1a41-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d1a41-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d1a41-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d1a41-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-713">- BindingEvents</span></span><br><span data-ttu-id="d1a41-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-714">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-716">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-717">
         - File</span></span><br><span data-ttu-id="d1a41-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-719">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-722">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-723">
         - Selection</span></span><br><span data-ttu-id="d1a41-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-724">
         - Settings</span></span><br><span data-ttu-id="d1a41-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-725">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-726">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-727">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-728">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-730">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-730">Office 2016 on Mac</span></span><br><span data-ttu-id="d1a41-731">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-732">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d1a41-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d1a41-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d1a41-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-736">- BindingEvents</span></span><br><span data-ttu-id="d1a41-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-737">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d1a41-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="d1a41-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-739">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-740">
         - File</span></span><br><span data-ttu-id="d1a41-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-742">
         - MatrixBindings</span></span><br><span data-ttu-id="d1a41-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="d1a41-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d1a41-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-745">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-746">
         - Selection</span></span><br><span data-ttu-id="d1a41-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-747">
         - Settings</span></span><br><span data-ttu-id="d1a41-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-748">
         - TableBindings</span></span><br><span data-ttu-id="d1a41-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-749">
         - TableCoercion</span></span><br><span data-ttu-id="d1a41-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d1a41-750">
         - TextBindings</span></span><br><span data-ttu-id="d1a41-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-751">
         - TextCoercion</span></span><br><span data-ttu-id="d1a41-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d1a41-753">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d1a41-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d1a41-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d1a41-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d1a41-755">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d1a41-755">Platform</span></span></th>
    <th><span data-ttu-id="d1a41-756">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-756">Extension points</span></span></th>
    <th><span data-ttu-id="d1a41-757">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="d1a41-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-759">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="d1a41-760">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-760">- Content</span></span><br><span data-ttu-id="d1a41-761">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-761">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d1a41-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d1a41-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-767">- ActiveView</span></span><br><span data-ttu-id="d1a41-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-768">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-769">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-770">
         - File</span></span><br><span data-ttu-id="d1a41-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-771">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-772">
         - Selection</span></span><br><span data-ttu-id="d1a41-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-773">
         - Settings</span></span><br><span data-ttu-id="d1a41-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-775">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-775">Office on Windows</span></span><br><span data-ttu-id="d1a41-776">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-777">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-777">- Content</span></span><br><span data-ttu-id="d1a41-778">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-778">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d1a41-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d1a41-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-784">- ActiveView</span></span><br><span data-ttu-id="d1a41-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-785">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-786">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-787">
         - File</span></span><br><span data-ttu-id="d1a41-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-788">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-789">
         - Selection</span></span><br><span data-ttu-id="d1a41-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-790">
         - Settings</span></span><br><span data-ttu-id="d1a41-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-792">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-792">Office 2019 on Windows</span></span><br><span data-ttu-id="d1a41-793">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-794">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-794">- Content</span></span><br><span data-ttu-id="d1a41-795">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-795">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-799">- ActiveView</span></span><br><span data-ttu-id="d1a41-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-800">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-801">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-802">
         - File</span></span><br><span data-ttu-id="d1a41-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-803">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-804">
         - Selection</span></span><br><span data-ttu-id="d1a41-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-805">
         - Settings</span></span><br><span data-ttu-id="d1a41-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-807">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-807">Office 2016 on Windows</span></span><br><span data-ttu-id="d1a41-808">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-809">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-809">- Content</span></span><br><span data-ttu-id="d1a41-810">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d1a41-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d1a41-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-813">- ActiveView</span></span><br><span data-ttu-id="d1a41-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-814">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-815">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-816">
         - File</span></span><br><span data-ttu-id="d1a41-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-817">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-818">
         - Selection</span></span><br><span data-ttu-id="d1a41-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-819">
         - Settings</span></span><br><span data-ttu-id="d1a41-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-821">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-821">Office 2013 on Windows</span></span><br><span data-ttu-id="d1a41-822">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-823">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-823">- Content</span></span><br><span data-ttu-id="d1a41-824">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d1a41-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d1a41-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d1a41-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-827">- ActiveView</span></span><br><span data-ttu-id="d1a41-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-828">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-829">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-830">
         - File</span></span><br><span data-ttu-id="d1a41-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-831">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-832">
         - Selection</span></span><br><span data-ttu-id="d1a41-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-833">
         - Settings</span></span><br><span data-ttu-id="d1a41-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-835">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d1a41-835">Office on iPad</span></span><br><span data-ttu-id="d1a41-836">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-837">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-837">- Content</span></span><br><span data-ttu-id="d1a41-838">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d1a41-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-842">- ActiveView</span></span><br><span data-ttu-id="d1a41-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-843">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-844">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-845">
         - File</span></span><br><span data-ttu-id="d1a41-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-846">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-847">
         - Selection</span></span><br><span data-ttu-id="d1a41-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-848">
         - Settings</span></span><br><span data-ttu-id="d1a41-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-850">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-850">Office on Mac</span></span><br><span data-ttu-id="d1a41-851">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1a41-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d1a41-852">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-852">- Content</span></span><br><span data-ttu-id="d1a41-853">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-853">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d1a41-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d1a41-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d1a41-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-859">- ActiveView</span></span><br><span data-ttu-id="d1a41-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-860">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-861">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-862">
         - File</span></span><br><span data-ttu-id="d1a41-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-863">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-864">
         - Selection</span></span><br><span data-ttu-id="d1a41-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-865">
         - Settings</span></span><br><span data-ttu-id="d1a41-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-867">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-867">Office 2019 on Mac</span></span><br><span data-ttu-id="d1a41-868">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-869">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-869">- Content</span></span><br><span data-ttu-id="d1a41-870">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-870">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-874">- ActiveView</span></span><br><span data-ttu-id="d1a41-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-875">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-876">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-877">
         - File</span></span><br><span data-ttu-id="d1a41-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-878">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-879">
         - Selection</span></span><br><span data-ttu-id="d1a41-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-880">
         - Settings</span></span><br><span data-ttu-id="d1a41-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-882">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-882">Office 2016 on Mac</span></span><br><span data-ttu-id="d1a41-883">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-884">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-884">- Content</span></span><br><span data-ttu-id="d1a41-885">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d1a41-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d1a41-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d1a41-888">- ActiveView</span></span><br><span data-ttu-id="d1a41-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-889">
         - CompressedFile</span></span><br><span data-ttu-id="d1a41-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-890">
         - DocumentEvents</span></span><br><span data-ttu-id="d1a41-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="d1a41-891">
         - File</span></span><br><span data-ttu-id="d1a41-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d1a41-892">
         - PdfFile</span></span><br><span data-ttu-id="d1a41-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-893">
         - Selection</span></span><br><span data-ttu-id="d1a41-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-894">
         - Settings</span></span><br><span data-ttu-id="d1a41-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d1a41-896">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d1a41-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d1a41-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="d1a41-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d1a41-898">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d1a41-898">Platform</span></span></th>
    <th><span data-ttu-id="d1a41-899">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-899">Extension points</span></span></th>
    <th><span data-ttu-id="d1a41-900">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="d1a41-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-902">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d1a41-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="d1a41-903">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d1a41-903">- Content</span></span><br><span data-ttu-id="d1a41-904">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-904">
         - TaskPane</span></span><br><span data-ttu-id="d1a41-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d1a41-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d1a41-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d1a41-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d1a41-909">- DocumentEvents</span></span><br><span data-ttu-id="d1a41-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="d1a41-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d1a41-911">
         - Settings</span></span><br><span data-ttu-id="d1a41-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d1a41-913">Projet</span><span class="sxs-lookup"><span data-stu-id="d1a41-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d1a41-914">Plateforme</span><span class="sxs-lookup"><span data-stu-id="d1a41-914">Platform</span></span></th>
    <th><span data-ttu-id="d1a41-915">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d1a41-915">Extension points</span></span></th>
    <th><span data-ttu-id="d1a41-916">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="d1a41-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d1a41-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-918">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-918">Office 2019 on Windows</span></span><br><span data-ttu-id="d1a41-919">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-920">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-922">- Selection</span></span><br><span data-ttu-id="d1a41-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-924">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-924">Office 2016 on Windows</span></span><br><span data-ttu-id="d1a41-925">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-926">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-928">- Selection</span></span><br><span data-ttu-id="d1a41-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d1a41-930">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d1a41-930">Office 2013 on Windows</span></span><br><span data-ttu-id="d1a41-931">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d1a41-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d1a41-932">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d1a41-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d1a41-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d1a41-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="d1a41-934">- Selection</span></span><br><span data-ttu-id="d1a41-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d1a41-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d1a41-936">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d1a41-936">See also</span></span>

- [<span data-ttu-id="d1a41-937">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d1a41-938">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d1a41-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d1a41-939">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="d1a41-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="d1a41-940">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="d1a41-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="d1a41-941">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="d1a41-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="d1a41-942">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="d1a41-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="d1a41-943">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="d1a41-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="d1a41-944">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="d1a41-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="d1a41-945">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d1a41-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="d1a41-946">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d1a41-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="d1a41-947">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="d1a41-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="d1a41-948">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="d1a41-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)