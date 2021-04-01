---
title: Application cliente Office et disponibilité de la plate-forme pour les compléments Office
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 02/19/2021
localization_priority: Priority
ms.openlocfilehash: e647142774ecec82224b52aa63c77103fc8dfe23
ms.sourcegitcommit: 85b4839be743059bf155ff44e49d64968444d80a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2021
ms.locfileid: "51471730"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="81e1a-103">Application cliente Office et disponibilité de la plate-forme pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="81e1a-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'une application Office spécifique, d'un ensemble de conditions requises, d'un membre de l’API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API courantes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="81e1a-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span data-ttu-id="81e1a-106"><span>Excel</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-106"><span>Excel</span></a></span></span>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span data-ttu-id="81e1a-107"><span>OneNote</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-107"><span>OneNote</span></a></span></span>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span data-ttu-id="81e1a-108"><span>Outlook</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-108"><span>Outlook</span></a></span></span>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span data-ttu-id="81e1a-109"><span>PowerPoint</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-109"><span>PowerPoint</span></a></span></span>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span data-ttu-id="81e1a-110"><span>Project</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-110"><span>Project</span></a></span></span>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span data-ttu-id="81e1a-111"><span>Word</span></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-111"><span>Word</span></a></span></span>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> <span data-ttu-id="81e1a-112">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="81e1a-112">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="81e1a-113">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="81e1a-113">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="81e1a-114">Excel</span><span class="sxs-lookup"><span data-stu-id="81e1a-114">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="81e1a-115">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="81e1a-115">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="81e1a-116">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-116">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="81e1a-117">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-117">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="81e1a-118"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-118"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-119">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-119">Office on the web</span></span></td>
    <td><span data-ttu-id="81e1a-120">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-120">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-121">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-121">
      - Content</span></span><br><span data-ttu-id="81e1a-122">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-122">
      - CustomFunctions</span></span><br><span data-ttu-id="81e1a-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="81e1a-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="81e1a-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="81e1a-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span></span><br><span data-ttu-id="81e1a-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="81e1a-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-143">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-143">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-144">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-144">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-145">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-145">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-146">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-146">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-147">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-147">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-151">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-151">Office on Windows</span></span><br><span data-ttu-id="81e1a-152">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-152">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-153">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-153">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-154">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-154">
      - Content</span></span><br><span data-ttu-id="81e1a-155">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-155">
      - CustomFunctions</span></span><br><span data-ttu-id="81e1a-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="81e1a-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="81e1a-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="81e1a-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span></span><br><span data-ttu-id="81e1a-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-175">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-175">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-176">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-176">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-177">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-177">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-178">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-178">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-185">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-185">Office 2019 on Windows</span></span><br><span data-ttu-id="81e1a-186">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-186">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-187">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-187">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-188">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-188">
      - Content</span></span><br><span data-ttu-id="81e1a-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-202">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-202">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-203">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-203">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-204">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-204">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-205">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-205">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-212">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-212">Office 2016 on Windows</span></span><br><span data-ttu-id="81e1a-213">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-213">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-214">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-214">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-215">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-215">
      - Content</span></span> </td>
    <td><span data-ttu-id="81e1a-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-221">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-221">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-222">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-222">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-223">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-223">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-224">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-224">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-231">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-231">Office 2013 on Windows</span></span><br><span data-ttu-id="81e1a-232">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-232">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-233">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-233">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-234">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-234">
      - Content</span></span> </td>
    <td><span data-ttu-id="81e1a-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-238">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-238">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-239">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-239">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-240">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-240">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-241">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-241">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-248">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="81e1a-248">Office on iPad</span></span><br><span data-ttu-id="81e1a-249">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-249">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-250">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-250">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-251">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-251">
      - Content</span></span> </td>
    <td><span data-ttu-id="81e1a-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="81e1a-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="81e1a-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="81e1a-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span></span><br><span data-ttu-id="81e1a-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-266">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-266">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-267">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-267">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-268">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-268">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-269">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-269">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-270">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-270">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-277">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-277">Office on Mac</span></span><br><span data-ttu-id="81e1a-278">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-278">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-279">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-279">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-280">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-280">
      - Content</span></span><br><span data-ttu-id="81e1a-281">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-281">
      - CustomFunctions</span></span><br><span data-ttu-id="81e1a-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="81e1a-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="81e1a-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="81e1a-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a></span></span><br><span data-ttu-id="81e1a-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-299">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-299">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-300">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-300">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-301">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-301">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-302">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-302">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-312">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-312">Office 2019 on Mac</span></span><br><span data-ttu-id="81e1a-313">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-313">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-314">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-314">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-315">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-315">
      - Content</span></span><br><span data-ttu-id="81e1a-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="81e1a-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="81e1a-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="81e1a-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="81e1a-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="81e1a-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="81e1a-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="81e1a-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-327">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-327">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-328">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-328">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-329">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-329">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-330">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-330">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-340">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-340">Office 2016 on Mac</span></span><br><span data-ttu-id="81e1a-341">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-341">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-342">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-342">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-343">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-343">
      - Content</span></span> </td>
    <td><span data-ttu-id="81e1a-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="81e1a-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-347">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-347">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-348">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-348">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-349">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-349">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-350">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-350">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-351">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-351">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-352">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-352">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-353">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-353">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-354">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-354">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-355">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-355">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-358">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-358">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-359">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-359">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="81e1a-360">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="81e1a-360">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="81e1a-361">Fonctions personnalisées (Excel seulement)</span><span class="sxs-lookup"><span data-stu-id="81e1a-361">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-362">Plateforme</span><span class="sxs-lookup"><span data-stu-id="81e1a-362">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-363">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-366">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-366">Office on the web</span></span></td>
    <td><span data-ttu-id="81e1a-367">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-367">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="81e1a-368">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-368">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="81e1a-369">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-369">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="81e1a-370">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-370">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-371">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-371">Office on Windows</span></span><br><span data-ttu-id="81e1a-372">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-372">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-373">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-373">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="81e1a-374">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-374">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="81e1a-375">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-375">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="81e1a-376">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-376">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-377">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-377">Office on Mac</span></span><br><span data-ttu-id="81e1a-378">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-378">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-379">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="81e1a-379">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="81e1a-380">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-380">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="81e1a-381">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-381">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="81e1a-382">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-382">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="81e1a-383">Outlook</span><span class="sxs-lookup"><span data-stu-id="81e1a-383">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-384">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="81e1a-384">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-385">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-385">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-386">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-386">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-387"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-387"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-388">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-388">Office on the web</span></span><br><span data-ttu-id="81e1a-389">(moderne)</span><span class="sxs-lookup"><span data-stu-id="81e1a-389">(modern)</span></span></td>
    <td><span data-ttu-id="81e1a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-392">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-392">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-393">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-393">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-394">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-394">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="81e1a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="81e1a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="81e1a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a></span></span><br><span data-ttu-id="81e1a-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="81e1a-405">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-405">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-406">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-406">Office on the web</span></span><br><span data-ttu-id="81e1a-407">(classique)</span><span class="sxs-lookup"><span data-stu-id="81e1a-407">(classic)</span></span></td>
    <td><span data-ttu-id="81e1a-408">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-408">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-409">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-409">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-410">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-410">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-411">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-411">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-412">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-412">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-419">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-419">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-420">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-420">Office on Windows</span></span><br><span data-ttu-id="81e1a-421">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-421">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-425">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-425">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="81e1a-427">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-427">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="81e1a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="81e1a-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="81e1a-436">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-436">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a></span></span><br><span data-ttu-id="81e1a-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="81e1a-438">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-438">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-439">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-439">Office 2019 on Windows</span></span><br><span data-ttu-id="81e1a-440">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-440">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-441">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-441">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-442">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-442">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-443">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-443">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-444">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-444">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-445">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-445">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="81e1a-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="81e1a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-454">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-455">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-455">Office 2016 on Windows</span></span><br><span data-ttu-id="81e1a-456">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-456">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-457">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-457">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-459">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-459">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-460">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-460">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="81e1a-462">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-462">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="81e1a-467">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-467">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-468">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-468">Office 2013 on Windows</span></span><br><span data-ttu-id="81e1a-469">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-469">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-473">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-473">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="81e1a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><sup>2</sup></span><span class="sxs-lookup"><span data-stu-id="81e1a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><sup>2</sup></span></span><br><span data-ttu-id="81e1a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="81e1a-478">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-479">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="81e1a-479">Office on iOS</span></span><br><span data-ttu-id="81e1a-480">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-480">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-481">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-481">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="81e1a-483">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-483">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-489">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-489">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-490">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-490">Office on Mac</span></span><br><span data-ttu-id="81e1a-491">(interface utilisateur actuelle</span><span class="sxs-lookup"><span data-stu-id="81e1a-491">(current UI,</span></span><br><span data-ttu-id="81e1a-492">connectée à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-492">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-493">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-493">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-494">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-494">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-495">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-495">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-497">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-497">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-498">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-498">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-499">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-499">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-500">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-500">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="81e1a-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="81e1a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="81e1a-506">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-506">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="81e1a-507">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-507">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-508">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-508">Office on Mac</span></span><br><span data-ttu-id="81e1a-509">(nouvelle interface utilisateur (aperçu)<sup>3</sup>,</span><span class="sxs-lookup"><span data-stu-id="81e1a-509">(new UI (preview)<sup>3</sup>,</span></span><br><span data-ttu-id="81e1a-510">connectée à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-510">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-511">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-511">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="81e1a-522">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-522">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-523">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-523">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-524">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-524">Office 2019 on Mac</span></span><br><span data-ttu-id="81e1a-525">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-525">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-526">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-526">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-527">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-527">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-528">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-528">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-529">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-529">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-530">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-530">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-531">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-531">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-532">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-532">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-533">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-533">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-534">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-534">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-535">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-535">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-536">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-536">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-537">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-537">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-538">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-538">Office 2016 on Mac</span></span><br><span data-ttu-id="81e1a-539">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-539">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-540">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-540">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-541">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-541">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="81e1a-542">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-542">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="81e1a-543">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-543">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="81e1a-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-545">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-545">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-546">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-546">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-547">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-547">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-548">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-548">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-549">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-549">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="81e1a-550">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-550">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-551">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-551">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-552">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="81e1a-552">Office on Android</span></span><br><span data-ttu-id="81e1a-553">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-553">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-554">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-554">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="81e1a-555">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-555">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="81e1a-556">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-556">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-557">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-557">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="81e1a-558">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-558">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="81e1a-559">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-559">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="81e1a-560">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-560">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="81e1a-561">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-561">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-562">Non disponible</span><span class="sxs-lookup"><span data-stu-id="81e1a-562">Not available</span></span></td>
  </tr>
</table>

> [!NOTE]
> <span data-ttu-id="81e1a-563"><sup>1</sup> Pour nécessiter le jeu d'API d'identité 1.3 dans votre code additionnel, vérifiez s'il est pris en charge en appelant `isSetSupported('IdentityAPI', '1.3')`.</span><span class="sxs-lookup"><span data-stu-id="81e1a-563"><sup>1</sup> To require Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="81e1a-564">Le déclarer dans le manifeste de votre macro complémentaire n'est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="81e1a-564">Declaring it in your add-in's manifest isn't supported.</span></span> <span data-ttu-id="81e1a-565">Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`.</span><span class="sxs-lookup"><span data-stu-id="81e1a-565">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="81e1a-566">Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="81e1a-566">For further details, see [Using APIs from later requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>
>
> <span data-ttu-id="81e1a-567"><sup>2</sup> Ajouté avec les mises à jour après la publication.</span><span class="sxs-lookup"><span data-stu-id="81e1a-567"><sup>2</sup> Added with post-release updates.</span></span>
>
> <span data-ttu-id="81e1a-568"><sup>3</sup> prise en charge de la préversion pour le nouvel Outlook sur Mac est disponible dans la version 16.38.506.</span><span class="sxs-lookup"><span data-stu-id="81e1a-568"><sup>3</sup> Support for the new Mac UI (preview) is available from Outlook version 16.38.506.</span></span> <span data-ttu-id="81e1a-569">Pour plus d’informations, consultez la section [Prise en charge du macro complémentaire dans Outlook sur le nouvel interface d’utilisateur Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).</span><span class="sxs-lookup"><span data-stu-id="81e1a-569">For more information, see the [Add-in support in Outlook on new Mac UI](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) section.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="81e1a-570">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="81e1a-570">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="81e1a-571">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="81e1a-571">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="81e1a-572">Word</span><span class="sxs-lookup"><span data-stu-id="81e1a-572">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-573">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="81e1a-573">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-574">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-574">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-575">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-575">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-576"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-576"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-577">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-577">Office on the web</span></span></td>
    <td><span data-ttu-id="81e1a-578">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-578">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-583">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-583">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-602">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-602">Office on Windows</span></span><br><span data-ttu-id="81e1a-603">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-603">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-604">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-604">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-609">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-609">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-630">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-630">Office 2019 on Windows</span></span><br><span data-ttu-id="81e1a-631">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-631">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-632">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-632">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-654">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-654">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-656">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-656">Office 2016 on Windows</span></span><br><span data-ttu-id="81e1a-657">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-657">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-658">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-658">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-679">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-679">Office 2013 on Windows</span></span><br><span data-ttu-id="81e1a-680">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-680">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-681">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-681">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-701">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="81e1a-701">Office on iPad</span></span><br><span data-ttu-id="81e1a-702">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-702">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-703">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-706">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-706">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-726">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-726">Office on Mac</span></span><br><span data-ttu-id="81e1a-727">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-727">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-728">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-728">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-732">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-732">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-754">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-754">Office 2019 on Mac</span></span><br><span data-ttu-id="81e1a-755">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-755">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-756">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-756">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-757">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-757">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-758">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-758">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-759">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-759">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="81e1a-760">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-760">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="81e1a-761">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-761">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-775">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-775">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-777">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-777">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-778">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-778">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-780">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-780">Office 2016 on Mac</span></span><br><span data-ttu-id="81e1a-781">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-781">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-782">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-782">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="81e1a-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="81e1a-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="81e1a-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-792">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-792">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="81e1a-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="81e1a-794">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-794">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="81e1a-795">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-795">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="81e1a-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="81e1a-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="81e1a-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="81e1a-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="81e1a-803">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="81e1a-803">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="81e1a-804">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="81e1a-804">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-805">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="81e1a-805">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-806">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-806">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-807">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-807">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-808"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-808"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-809">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-809">Office on the web</span></span></td>
    <td><span data-ttu-id="81e1a-810">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-810">
      - Content</span></span><br><span data-ttu-id="81e1a-811">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-811">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="81e1a-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="81e1a-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-821">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-821">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-822">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-822">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-823">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-823">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-824">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-824">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-827">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-827">Office on Windows</span></span><br><span data-ttu-id="81e1a-828">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-828">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-829">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-829">
      - Content</span></span><br><span data-ttu-id="81e1a-830">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-830">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="81e1a-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="81e1a-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-835">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-835">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-836">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-836">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-837">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-837">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-838">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-838">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-846">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-846">Office 2019 on Windows</span></span><br><span data-ttu-id="81e1a-847">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-847">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-848">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-848">
      - Content</span></span><br><span data-ttu-id="81e1a-849">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-849">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-850">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-850">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-851">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-851">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-852">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-852">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-853">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-853">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-861">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-861">Office 2016 on Windows</span></span><br><span data-ttu-id="81e1a-862">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-862">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-863">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-863">
      - Content</span></span><br><span data-ttu-id="81e1a-864">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-864">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="81e1a-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-867">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-867">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-869">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-869">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-870">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-870">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-875">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-875">Office 2013 on Windows</span></span><br><span data-ttu-id="81e1a-876">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-876">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-877">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-877">
      - Content</span></span><br><span data-ttu-id="81e1a-878">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-878">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="81e1a-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-882">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-882">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-883">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-883">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-884">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-884">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-885">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-885">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-886">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-886">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-889">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="81e1a-889">Office on iPad</span></span><br><span data-ttu-id="81e1a-890">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-890">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-891">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-891">
      - Content</span></span><br><span data-ttu-id="81e1a-892">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-892">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="81e1a-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="81e1a-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="81e1a-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-896">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-896">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-897">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-897">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-898">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-898">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-899">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-899">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-900">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-900">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-901">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-901">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-902">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-902">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-903">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-903">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-904">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-904">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-905">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-905">Office on Mac</span></span><br><span data-ttu-id="81e1a-906">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="81e1a-906">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="81e1a-907">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-907">
      - Content</span></span><br><span data-ttu-id="81e1a-908">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-908">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="81e1a-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="81e1a-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-913">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-913">
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="81e1a-914">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-914">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="81e1a-915">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-915">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-916">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-916">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-917">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-917">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-918">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-918">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-919">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-919">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-920">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-920">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-921">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-921">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-924">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-924">Office 2019 on Mac</span></span><br><span data-ttu-id="81e1a-925">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-925">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-926">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-926">
      - Content</span></span><br><span data-ttu-id="81e1a-927">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-927">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-930">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-930">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-931">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-931">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-932">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-932">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-933">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-933">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-938">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-938">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-939">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-939">Office 2016 on Mac</span></span><br><span data-ttu-id="81e1a-940">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-940">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-941">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-941">
      - Content</span></span><br><span data-ttu-id="81e1a-942">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-942">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="81e1a-943">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="81e1a-943">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="81e1a-944">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-944">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-945">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-945">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="81e1a-946">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-946">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="81e1a-947">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-947">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-948">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-948">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="81e1a-949">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-949">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="81e1a-950">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-950">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-951">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-951">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-952">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-952">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="81e1a-953">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="81e1a-953">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="81e1a-954">OneNote</span><span class="sxs-lookup"><span data-stu-id="81e1a-954">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-955">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="81e1a-955">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-956">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-956">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-957">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-957">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-958"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-958"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-959">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="81e1a-959">Office on the web</span></span></td>
    <td><span data-ttu-id="81e1a-960">
      - Contenu</span><span class="sxs-lookup"><span data-stu-id="81e1a-960">
      - Content</span></span><br><span data-ttu-id="81e1a-961">
      - Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-961">
      - TaskPane</span></span><br><span data-ttu-id="81e1a-962">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-962">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-963">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-963">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="81e1a-964">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-964">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="81e1a-965">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-965">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="81e1a-966">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-966">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="81e1a-967">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-967">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="81e1a-968">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Paramètres</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-968">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="81e1a-969">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-969">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="81e1a-970">Project</span><span class="sxs-lookup"><span data-stu-id="81e1a-970">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="81e1a-971">Plateforme</span><span class="sxs-lookup"><span data-stu-id="81e1a-971">Platform</span></span></th>
    <th><span data-ttu-id="81e1a-972">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="81e1a-972">Extension points</span></span></th>
    <th><span data-ttu-id="81e1a-973">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-973">API requirement sets</span></span></th>
    <th><span data-ttu-id="81e1a-974"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="81e1a-974"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-975">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-975">Office 2019 on Windows</span></span><br><span data-ttu-id="81e1a-976">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-976">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-977">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-977">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-978">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-978">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="81e1a-979">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-979">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-980">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-980">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-981">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-981">Office 2016 on Windows</span></span><br><span data-ttu-id="81e1a-982">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-982">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-983">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-983">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-984">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-984">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="81e1a-985">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-985">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-986">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-986">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="81e1a-987">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="81e1a-987">Office 2013 on Windows</span></span><br><span data-ttu-id="81e1a-988">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="81e1a-988">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="81e1a-989">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-989">- TaskPane</span></span></td>
    <td><span data-ttu-id="81e1a-990">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-990">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="81e1a-991">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="81e1a-991">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="81e1a-992">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="81e1a-992">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="81e1a-993">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="81e1a-993">See also</span></span>

- [<span data-ttu-id="81e1a-994">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-994">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="81e1a-995">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="81e1a-995">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="81e1a-996">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="81e1a-996">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="81e1a-997">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="81e1a-997">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="81e1a-998">Documentation de référence de l’API</span><span class="sxs-lookup"><span data-stu-id="81e1a-998">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="81e1a-999">Historique des mises à jour de Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="81e1a-999">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="81e1a-1000">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="81e1a-1000">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="81e1a-1001">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="81e1a-1001">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="81e1a-1002">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="81e1a-1002">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="81e1a-1003">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="81e1a-1003">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="81e1a-1004">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="81e1a-1004">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="81e1a-1005">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="81e1a-1005">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
