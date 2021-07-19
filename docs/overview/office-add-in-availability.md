---
title: Application cliente Office et disponibilité de la plate-forme pour les compléments Office
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 7b3bd770d74f29d1a0b27da5080284aa62146101
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455494"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="8ab20-103">Application cliente Office et disponibilité de la plate-forme pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="8ab20-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'une application Office spécifique, d'un ensemble de conditions requises, d'un membre de l’API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API courantes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="8ab20-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span data-ttu-id="8ab20-106"><span>Excel</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-106"><span>Excel</span></a></span></span>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span data-ttu-id="8ab20-107"><span>OneNote</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-107"><span>OneNote</span></a></span></span>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span data-ttu-id="8ab20-108"><span>Outlook</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-108"><span>Outlook</span></a></span></span>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span data-ttu-id="8ab20-109"><span>PowerPoint</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-109"><span>PowerPoint</span></a></span></span>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span data-ttu-id="8ab20-110"><span>Project</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-110"><span>Project</span></a></span></span>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span data-ttu-id="8ab20-111"><span>Word</span></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-111"><span>Word</span></a></span></span>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> <span data-ttu-id="8ab20-112">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="8ab20-112">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="8ab20-113">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="8ab20-113">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span> <span data-ttu-id="8ab20-114">Les compléments Office ne sont peut-être pas pris en charge sur tous les services membres du [Programme partenaire de stockage cloud Office](https://developer.microsoft.com/office/cloud-storage-partner-program), qui permet à l’intégration d’Office sur le web d’utiliser des documents Office dans le cadre de leur offre de service.</span><span class="sxs-lookup"><span data-stu-id="8ab20-114">Office Add-ins may not be supported on all services that are members of the [Office Cloud Storage Partner Program](https://developer.microsoft.com/office/cloud-storage-partner-program), which enables integrating Office on the web to work with Office documents as part of their service offering.</span></span> <span data-ttu-id="8ab20-115">Pour plus d’informations, contactez le service membre.</span><span class="sxs-lookup"><span data-stu-id="8ab20-115">For more information, please contact the member service.</span></span>

## <a name="excel"></a><span data-ttu-id="8ab20-116">Excel</span><span class="sxs-lookup"><span data-stu-id="8ab20-116">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8ab20-117">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8ab20-117">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8ab20-118">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-118">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8ab20-119">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-119">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8ab20-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-121">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-121">Office on the web</span></span></td>
    <td><span data-ttu-id="8ab20-122">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-122">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-123">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-123">
      - Content</span></span><br><span data-ttu-id="8ab20-124">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-124">
      - CustomFunctions</span></span><br><span data-ttu-id="8ab20-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8ab20-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8ab20-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="8ab20-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="8ab20-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="8ab20-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span></span><br><span data-ttu-id="8ab20-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="8ab20-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-156">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-156">Office on Windows</span></span><br><span data-ttu-id="8ab20-157">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-157">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-158">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-158">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-159">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-159">
      - Content</span></span><br><span data-ttu-id="8ab20-160">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-160">
      - CustomFunctions</span></span><br><span data-ttu-id="8ab20-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8ab20-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8ab20-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="8ab20-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="8ab20-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="8ab20-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="8ab20-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="8ab20-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-194">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-194">Office 2019 on Windows</span></span><br><span data-ttu-id="8ab20-195">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-195">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-196">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-196">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-197">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-197">
      - Content</span></span><br><span data-ttu-id="8ab20-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-221">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-221">Office 2016 on Windows</span></span><br><span data-ttu-id="8ab20-222">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-223">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-223">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-224">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="8ab20-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-240">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-240">Office 2013 on Windows</span></span><br><span data-ttu-id="8ab20-241">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-241">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-242">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-242">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-243">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-243">
      - Content</span></span> </td>
    <td><span data-ttu-id="8ab20-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-257">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="8ab20-257">Office on iPad</span></span><br><span data-ttu-id="8ab20-258">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-258">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-259">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-259">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-260">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-260">
      - Content</span></span> </td>
    <td><span data-ttu-id="8ab20-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8ab20-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8ab20-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="8ab20-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="8ab20-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="8ab20-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-288">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-288">Office on Mac</span></span><br><span data-ttu-id="8ab20-289">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-289">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-290">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-290">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-291">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-291">
      - Content</span></span><br><span data-ttu-id="8ab20-292">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-292">
      - CustomFunctions</span></span><br><span data-ttu-id="8ab20-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8ab20-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8ab20-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="8ab20-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="8ab20-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="8ab20-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="8ab20-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="8ab20-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-327">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-327">Office 2019 on Mac</span></span><br><span data-ttu-id="8ab20-328">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-329">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-329">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-330">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-330">
      - Content</span></span><br><span data-ttu-id="8ab20-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8ab20-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8ab20-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8ab20-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8ab20-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8ab20-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8ab20-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8ab20-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-355">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-355">Office 2016 on Mac</span></span><br><span data-ttu-id="8ab20-356">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-356">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-357">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-357">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-358">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-358">
      - Content</span></span> </td>
    <td><span data-ttu-id="8ab20-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8ab20-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="8ab20-375">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="8ab20-375">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="8ab20-376">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="8ab20-376">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-377">Plateforme</span><span class="sxs-lookup"><span data-stu-id="8ab20-377">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-378">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-378">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-379">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-379">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-381">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-381">Office on the web</span></span></td>
    <td><span data-ttu-id="8ab20-382">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-382">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="8ab20-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="8ab20-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="8ab20-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-386">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-386">Office on Windows</span></span><br><span data-ttu-id="8ab20-387">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-387">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-388">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-388">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="8ab20-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="8ab20-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="8ab20-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-392">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-392">Office on Mac</span></span><br><span data-ttu-id="8ab20-393">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-393">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-394">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="8ab20-394">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="8ab20-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="8ab20-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="8ab20-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="8ab20-398">Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab20-398">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-399">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8ab20-399">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-400">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-400">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-401">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-401">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-403">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-403">Office on the web</span></span><br><span data-ttu-id="8ab20-404">(moderne)</span><span class="sxs-lookup"><span data-stu-id="8ab20-404">(modern)</span></span></td>
    <td><span data-ttu-id="8ab20-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="8ab20-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="8ab20-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="8ab20-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="8ab20-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="8ab20-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="8ab20-421">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-421">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-422">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-422">Office on the web</span></span><br><span data-ttu-id="8ab20-423">(classique)</span><span class="sxs-lookup"><span data-stu-id="8ab20-423">(classic)</span></span></td>
    <td><span data-ttu-id="8ab20-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-435">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-436">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-436">Office on Windows</span></span><br><span data-ttu-id="8ab20-437">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-437">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="8ab20-443">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-443">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="8ab20-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="8ab20-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="8ab20-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="8ab20-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="8ab20-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="8ab20-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="8ab20-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-457">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-457">Office 2019 on Windows</span></span><br><span data-ttu-id="8ab20-458">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-458">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="8ab20-464">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-464">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="8ab20-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-472">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-472">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-473">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-473">Office 2016 on Windows</span></span><br><span data-ttu-id="8ab20-474">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-474">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="8ab20-480">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-480">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="8ab20-485">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-485">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-486">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-486">Office 2013 on Windows</span></span><br><span data-ttu-id="8ab20-487">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-487">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="8ab20-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup></span><span class="sxs-lookup"><span data-stu-id="8ab20-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup></span></span><br><span data-ttu-id="8ab20-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="8ab20-496">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-497">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="8ab20-497">Office on iOS</span></span><br><span data-ttu-id="8ab20-498">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-498">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="8ab20-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-507">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-507">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-508">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-508">Office on Mac</span></span><br><span data-ttu-id="8ab20-509">(interface utilisateur actuelle</span><span class="sxs-lookup"><span data-stu-id="8ab20-509">(current UI,</span></span><br><span data-ttu-id="8ab20-510">connectée à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-510">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="8ab20-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="8ab20-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="8ab20-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="8ab20-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="8ab20-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-526">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-526">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-527">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-527">Office on Mac</span></span><br><span data-ttu-id="8ab20-528">(nouvelle interface utilisateur (aperçu)<sup>3</sup>,</span><span class="sxs-lookup"><span data-stu-id="8ab20-528">(new UI (preview)<sup>3</sup>,</span></span><br><span data-ttu-id="8ab20-529">connectée à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-529">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="8ab20-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="8ab20-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="8ab20-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="8ab20-544">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-544">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-545">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-545">Office 2019 on Mac</span></span><br><span data-ttu-id="8ab20-546">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-546">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-558">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-558">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-559">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-559">Office 2016 on Mac</span></span><br><span data-ttu-id="8ab20-560">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-560">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8ab20-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8ab20-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8ab20-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="8ab20-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-572">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-572">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-573">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="8ab20-573">Office on Android</span></span><br><span data-ttu-id="8ab20-574">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-574">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8ab20-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="8ab20-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="8ab20-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="8ab20-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="8ab20-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="8ab20-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-583">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8ab20-583">Not available</span></span></td>
  </tr>
</table>

> [!NOTE]
> <span data-ttu-id="8ab20-584"><sup>1</sup> Pour nécessiter le jeu d'API d'identité 1.3 dans votre code additionnel, vérifiez s'il est pris en charge en appelant `isSetSupported('IdentityAPI', '1.3')`.</span><span class="sxs-lookup"><span data-stu-id="8ab20-584"><sup>1</sup> To require Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="8ab20-585">Le déclarer dans le manifeste de votre macro complémentaire n'est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="8ab20-585">Declaring it in your add-in's manifest isn't supported.</span></span> <span data-ttu-id="8ab20-586">Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8ab20-586">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="8ab20-587">Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8ab20-587">For further details, see [Using APIs from later requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>
>
> <span data-ttu-id="8ab20-588"><sup>2</sup> Ajouté avec les mises à jour après la publication.</span><span class="sxs-lookup"><span data-stu-id="8ab20-588"><sup>2</sup> Added with post-release updates.</span></span>
>
> <span data-ttu-id="8ab20-589"><sup>3</sup> prise en charge de la préversion pour le nouvel Outlook sur Mac est disponible dans la version 16.38.506.</span><span class="sxs-lookup"><span data-stu-id="8ab20-589"><sup>3</sup> Support for the new Mac UI (preview) is available from Outlook version 16.38.506.</span></span> <span data-ttu-id="8ab20-590">Pour plus d’informations, consultez la section [Prise en charge du macro complémentaire dans Outlook sur le nouvel interface d’utilisateur Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).</span><span class="sxs-lookup"><span data-stu-id="8ab20-590">For more information, see the [Add-in support in Outlook on new Mac UI](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) section.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8ab20-591">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8ab20-591">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="8ab20-592">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="8ab20-592">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="8ab20-593">Word</span><span class="sxs-lookup"><span data-stu-id="8ab20-593">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-594">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8ab20-594">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-595">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-595">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-596">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-596">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-598">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-598">Office on the web</span></span></td>
    <td><span data-ttu-id="8ab20-599">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-599">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-623">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-623">Office on Windows</span></span><br><span data-ttu-id="8ab20-624">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-624">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-625">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-625">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-652">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-652">Office 2019 on Windows</span></span><br><span data-ttu-id="8ab20-653">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-653">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-654">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-654">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-678">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-678">Office 2016 on Windows</span></span><br><span data-ttu-id="8ab20-679">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-679">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-680">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-680">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-701">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-701">Office 2013 on Windows</span></span><br><span data-ttu-id="8ab20-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-702">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-703">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-723">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="8ab20-723">Office on iPad</span></span><br><span data-ttu-id="8ab20-724">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-724">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-725">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-725">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-748">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-748">Office on Mac</span></span><br><span data-ttu-id="8ab20-749">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-749">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-750">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-750">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-777">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-777">Office 2019 on Mac</span></span><br><span data-ttu-id="8ab20-778">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-778">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-779">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-779">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="8ab20-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="8ab20-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-803">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-803">Office 2016 on Mac</span></span><br><span data-ttu-id="8ab20-804">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-804">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-805">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-805">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="8ab20-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="8ab20-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="8ab20-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="8ab20-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="8ab20-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="8ab20-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="8ab20-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="8ab20-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="8ab20-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="8ab20-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="8ab20-826">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="8ab20-826">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="8ab20-827">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8ab20-827">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-828">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8ab20-828">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-829">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-829">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-830">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-830">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-832">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-832">Office on the web</span></span></td>
    <td><span data-ttu-id="8ab20-833">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-833">
      - Content</span></span><br><span data-ttu-id="8ab20-834">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-834">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8ab20-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="8ab20-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-850">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-850">Office on Windows</span></span><br><span data-ttu-id="8ab20-851">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-851">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-852">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-852">
      - Content</span></span><br><span data-ttu-id="8ab20-853">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-853">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8ab20-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="8ab20-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-870">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-870">Office 2019 on Windows</span></span><br><span data-ttu-id="8ab20-871">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-871">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-872">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-872">
      - Content</span></span><br><span data-ttu-id="8ab20-873">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-873">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-885">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-885">Office 2016 on Windows</span></span><br><span data-ttu-id="8ab20-886">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-886">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-887">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-887">
      - Content</span></span><br><span data-ttu-id="8ab20-888">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-888">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="8ab20-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-899">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-899">Office 2013 on Windows</span></span><br><span data-ttu-id="8ab20-900">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-900">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-901">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-901">
      - Content</span></span><br><span data-ttu-id="8ab20-902">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-902">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="8ab20-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-913">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="8ab20-913">Office on iPad</span></span><br><span data-ttu-id="8ab20-914">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-914">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-915">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-915">
      - Content</span></span><br><span data-ttu-id="8ab20-916">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-916">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="8ab20-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8ab20-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="8ab20-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-930">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-930">Office on Mac</span></span><br><span data-ttu-id="8ab20-931">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="8ab20-931">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="8ab20-932">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-932">
      - Content</span></span><br><span data-ttu-id="8ab20-933">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-933">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8ab20-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="8ab20-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="8ab20-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8ab20-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="8ab20-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-950">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-950">Office 2019 on Mac</span></span><br><span data-ttu-id="8ab20-951">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-951">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-952">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-952">
      - Content</span></span><br><span data-ttu-id="8ab20-953">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-953">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-965">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-965">Office 2016 on Mac</span></span><br><span data-ttu-id="8ab20-966">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-966">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-967">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-967">
      - Content</span></span><br><span data-ttu-id="8ab20-968">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-968">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="8ab20-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8ab20-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8ab20-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="8ab20-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="8ab20-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="8ab20-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="8ab20-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="8ab20-979">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="8ab20-979">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="8ab20-980">OneNote</span><span class="sxs-lookup"><span data-stu-id="8ab20-980">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-981">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8ab20-981">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-982">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-982">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-983">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-983">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-985">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="8ab20-985">Office on the web</span></span></td>
    <td><span data-ttu-id="8ab20-986">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="8ab20-986">
      - Content</span></span><br><span data-ttu-id="8ab20-987">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-987">
      - TaskPane</span></span><br><span data-ttu-id="8ab20-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="8ab20-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="8ab20-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="8ab20-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="8ab20-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="8ab20-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="8ab20-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="8ab20-996">Project</span><span class="sxs-lookup"><span data-stu-id="8ab20-996">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8ab20-997">Plateforme</span><span class="sxs-lookup"><span data-stu-id="8ab20-997">Platform</span></span></th>
    <th><span data-ttu-id="8ab20-998">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8ab20-998">Extension points</span></span></th>
    <th><span data-ttu-id="8ab20-999">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-999">API requirement sets</span></span></th>
    <th><span data-ttu-id="8ab20-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-1001">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-1001">Office 2019 on Windows</span></span><br><span data-ttu-id="8ab20-1002">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1002">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-1003">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-1003">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8ab20-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-1007">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-1007">Office 2016 on Windows</span></span><br><span data-ttu-id="8ab20-1008">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1008">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-1009">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-1009">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8ab20-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8ab20-1013">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="8ab20-1013">Office 2013 on Windows</span></span><br><span data-ttu-id="8ab20-1014">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1014">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8ab20-1015">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-1015">- TaskPane</span></span></td>
    <td><span data-ttu-id="8ab20-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8ab20-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="8ab20-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="8ab20-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="8ab20-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="8ab20-1019">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8ab20-1019">See also</span></span>

- [<span data-ttu-id="8ab20-1020">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-1020">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="8ab20-1021">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8ab20-1021">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="8ab20-1022">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="8ab20-1022">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="8ab20-1023">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="8ab20-1023">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="8ab20-1024">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="8ab20-1024">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="8ab20-1025">Historique des mises à jour de Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="8ab20-1025">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="8ab20-1026">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1026">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="8ab20-1027">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1027">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="8ab20-1028">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1028">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="8ab20-1029">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="8ab20-1029">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="8ab20-1030">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="8ab20-1030">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="8ab20-1031">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="8ab20-1031">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
