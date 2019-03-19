---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 03/15/2019
localization_priority: Priority
ms.openlocfilehash: 4348881c35e4c79975d34406e4668b2693405134
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/16/2019
ms.locfileid: "30654962"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b6e33-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="b6e33-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b6e33-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="b6e33-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b6e33-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="b6e33-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="b6e33-108">Le numéro de build d’un achat définitif d’Office 2019 est 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="b6e33-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="b6e33-109">Excel</span><span class="sxs-lookup"><span data-stu-id="b6e33-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b6e33-110">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b6e33-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b6e33-111">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b6e33-112">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b6e33-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e33-114">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e33-115">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-115">- TaskPane</span></span><br><span data-ttu-id="b6e33-116">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-116">
        - Content</span></span><br><span data-ttu-id="b6e33-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b6e33-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b6e33-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-127">
        - BindingEvents</span></span><br><span data-ttu-id="b6e33-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-128">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-129">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-130">
        - File</span></span><br><span data-ttu-id="b6e33-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-131">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-133">
        - Selection</span></span><br><span data-ttu-id="b6e33-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-134">
        - Settings</span></span><br><span data-ttu-id="b6e33-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-135">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-136">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-137">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-139">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-140">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-140">- TaskPane</span></span><br><span data-ttu-id="b6e33-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-141">
        - Content</span></span><br><span data-ttu-id="b6e33-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b6e33-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b6e33-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-152">
        - BindingEvents</span></span><br><span data-ttu-id="b6e33-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-153">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-154">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-155">
        - File</span></span><br><span data-ttu-id="b6e33-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-156">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-158">
        - Selection</span></span><br><span data-ttu-id="b6e33-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-159">
        - Settings</span></span><br><span data-ttu-id="b6e33-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-160">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-161">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-162">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-164">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="b6e33-165">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-165">- TaskPane</span></span><br><span data-ttu-id="b6e33-166">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-166">
        - Content</span></span><br><span data-ttu-id="b6e33-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e33-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-177">- BindingEvents</span></span><br><span data-ttu-id="b6e33-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-178">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-179">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-180">
        - File</span></span><br><span data-ttu-id="b6e33-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-181">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-182">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-184">
        - Selection</span></span><br><span data-ttu-id="b6e33-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-185">
        - Settings</span></span><br><span data-ttu-id="b6e33-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-186">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-187">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-188">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-190">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="b6e33-191">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-191">- TaskPane</span></span><br><span data-ttu-id="b6e33-192">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-192">
        - Content</span></span></td>
    <td><span data-ttu-id="b6e33-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b6e33-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="b6e33-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-195">- BindingEvents</span></span><br><span data-ttu-id="b6e33-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-196">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-197">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-198">
        - File</span></span><br><span data-ttu-id="b6e33-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-199">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-200">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-202">
        - Selection</span></span><br><span data-ttu-id="b6e33-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-203">
        - Settings</span></span><br><span data-ttu-id="b6e33-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-204">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-205">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-206">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-208">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="b6e33-209">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-209">
        - TaskPane</span></span><br><span data-ttu-id="b6e33-210">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b6e33-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b6e33-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="b6e33-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-212">
        - BindingEvents</span></span><br><span data-ttu-id="b6e33-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-213">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-214">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-215">
        - File</span></span><br><span data-ttu-id="b6e33-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-216">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-217">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-219">
        - Selection</span></span><br><span data-ttu-id="b6e33-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-220">
        - Settings</span></span><br><span data-ttu-id="b6e33-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-221">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-222">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-223">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-225">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="b6e33-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="b6e33-226">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-226">- TaskPane</span></span><br><span data-ttu-id="b6e33-227">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-227">
        - Content</span></span></td>
    <td><span data-ttu-id="b6e33-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-237">- BindingEvents</span></span><br><span data-ttu-id="b6e33-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-238">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-239">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-240">
        - File</span></span><br><span data-ttu-id="b6e33-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-241">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-242">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-244">
        - Selection</span></span><br><span data-ttu-id="b6e33-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-245">
        - Settings</span></span><br><span data-ttu-id="b6e33-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-246">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-247">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-248">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-250">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="b6e33-251">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-251">- TaskPane</span></span><br><span data-ttu-id="b6e33-252">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-252">
        - Content</span></span><br><span data-ttu-id="b6e33-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e33-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-263">- BindingEvents</span></span><br><span data-ttu-id="b6e33-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-264">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-265">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-266">
        - File</span></span><br><span data-ttu-id="b6e33-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-267">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-268">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-270">
        - PdfFile</span></span><br><span data-ttu-id="b6e33-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-271">
        - Selection</span></span><br><span data-ttu-id="b6e33-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-272">
        - Settings</span></span><br><span data-ttu-id="b6e33-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-273">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-274">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-275">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-277">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="b6e33-278">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-278">- TaskPane</span></span><br><span data-ttu-id="b6e33-279">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-279">
        - Content</span></span><br><span data-ttu-id="b6e33-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e33-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e33-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e33-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e33-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e33-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e33-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b6e33-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e33-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e33-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-290">- BindingEvents</span></span><br><span data-ttu-id="b6e33-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-291">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-292">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-293">
        - File</span></span><br><span data-ttu-id="b6e33-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-294">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-295">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-297">
        - PdfFile</span></span><br><span data-ttu-id="b6e33-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-298">
        - Selection</span></span><br><span data-ttu-id="b6e33-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-299">
        - Settings</span></span><br><span data-ttu-id="b6e33-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-300">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-301">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-302">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-304">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="b6e33-305">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-305">- TaskPane</span></span><br><span data-ttu-id="b6e33-306">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-306">
        - Content</span></span></td>
    <td><span data-ttu-id="b6e33-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e33-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b6e33-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="b6e33-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-309">- BindingEvents</span></span><br><span data-ttu-id="b6e33-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-310">
        - CompressedFile</span></span><br><span data-ttu-id="b6e33-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-311">
        - DocumentEvents</span></span><br><span data-ttu-id="b6e33-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-312">
        - File</span></span><br><span data-ttu-id="b6e33-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-313">
        - ImageCoercion</span></span><br><span data-ttu-id="b6e33-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-314">
        - MatrixBindings</span></span><br><span data-ttu-id="b6e33-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-316">
        - PdfFile</span></span><br><span data-ttu-id="b6e33-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-317">
        - Selection</span></span><br><span data-ttu-id="b6e33-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-318">
        - Settings</span></span><br><span data-ttu-id="b6e33-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-319">
        - TableBindings</span></span><br><span data-ttu-id="b6e33-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-320">
        - TableCoercion</span></span><br><span data-ttu-id="b6e33-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-321">
        - TextBindings</span></span><br><span data-ttu-id="b6e33-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b6e33-323">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b6e33-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="b6e33-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="b6e33-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e33-325">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b6e33-325">Platform</span></span></th>
    <th><span data-ttu-id="b6e33-326">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-326">Extension points</span></span></th>
    <th><span data-ttu-id="b6e33-327">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e33-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e33-329">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e33-330">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-330">- Mail Read</span></span><br><span data-ttu-id="b6e33-331">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-331">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e33-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e33-340">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-341">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-342">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-342">- Mail Read</span></span><br><span data-ttu-id="b6e33-343">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-343">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b6e33-345">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b6e33-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b6e33-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e33-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e33-353">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-354">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-355">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-355">- Mail Read</span></span><br><span data-ttu-id="b6e33-356">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-356">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b6e33-358">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b6e33-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b6e33-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e33-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e33-366">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-367">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-368">- Mail Read</span></span><br><span data-ttu-id="b6e33-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-369">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b6e33-371">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b6e33-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b6e33-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b6e33-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-377">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-378">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-378">- Mail Read</span></span><br><span data-ttu-id="b6e33-379">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="b6e33-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b6e33-384">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-385">Office 365 pour iOS</span><span class="sxs-lookup"><span data-stu-id="b6e33-385">See the Office 365 SDK for iOS.</span></span></td>
    <td> <span data-ttu-id="b6e33-386">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-386">- Mail Read</span></span><br><span data-ttu-id="b6e33-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b6e33-393">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-394">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-395">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-395">- Mail Read</span></span><br><span data-ttu-id="b6e33-396">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-396">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b6e33-404">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-405">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-406">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-406">- Mail Read</span></span><br><span data-ttu-id="b6e33-407">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-407">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b6e33-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-416">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-417">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-417">- Mail Read</span></span><br><span data-ttu-id="b6e33-418">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-418">
      - Mail Compose</span></span><br><span data-ttu-id="b6e33-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e33-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b6e33-426">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-427">Office 365 pour Android</span><span class="sxs-lookup"><span data-stu-id="b6e33-427">See the Office 365 SDK for Android.</span></span></td>
    <td> <span data-ttu-id="b6e33-428">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b6e33-428">- Mail Read</span></span><br><span data-ttu-id="b6e33-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e33-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e33-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e33-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e33-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b6e33-435">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b6e33-435">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="b6e33-436">Word</span><span class="sxs-lookup"><span data-stu-id="b6e33-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e33-437">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b6e33-437">Platform</span></span></th>
    <th><span data-ttu-id="b6e33-438">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-438">Extension points</span></span></th>
    <th><span data-ttu-id="b6e33-439">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e33-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e33-441">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e33-442">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-442">- TaskPane</span></span><br><span data-ttu-id="b6e33-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-448">- BindingEvents</span></span><br><span data-ttu-id="b6e33-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-450">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-451">
         - File</span></span><br><span data-ttu-id="b6e33-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-453">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-454">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-457">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-458">
         - Selection</span></span><br><span data-ttu-id="b6e33-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-459">
         - Settings</span></span><br><span data-ttu-id="b6e33-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-460">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-461">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-462">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-463">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-465">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-466">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-466">- TaskPane</span></span><br><span data-ttu-id="b6e33-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-472">- BindingEvents</span></span><br><span data-ttu-id="b6e33-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-473">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-475">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-476">
         - File</span></span><br><span data-ttu-id="b6e33-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-478">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-479">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-482">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-483">
         - Selection</span></span><br><span data-ttu-id="b6e33-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-484">
         - Settings</span></span><br><span data-ttu-id="b6e33-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-485">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-486">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-487">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-488">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-490">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-491">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-491">- TaskPane</span></span><br><span data-ttu-id="b6e33-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-497">- BindingEvents</span></span><br><span data-ttu-id="b6e33-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-498">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-500">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-501">
         - File</span></span><br><span data-ttu-id="b6e33-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-503">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-504">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-507">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-508">
         - Selection</span></span><br><span data-ttu-id="b6e33-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-509">
         - Settings</span></span><br><span data-ttu-id="b6e33-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-510">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-511">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-512">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-513">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-515">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-516">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b6e33-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="b6e33-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-519">- BindingEvents</span></span><br><span data-ttu-id="b6e33-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-520">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-522">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-523">
         - File</span></span><br><span data-ttu-id="b6e33-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-525">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-526">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-529">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-530">
         - Selection</span></span><br><span data-ttu-id="b6e33-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-531">
         - Settings</span></span><br><span data-ttu-id="b6e33-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-532">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-533">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-534">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-535">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-537">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-538">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b6e33-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b6e33-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-540">- BindingEvents</span></span><br><span data-ttu-id="b6e33-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-541">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-543">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-544">
         - File</span></span><br><span data-ttu-id="b6e33-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-546">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-547">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-550">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-551">
         - Selection</span></span><br><span data-ttu-id="b6e33-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-552">
         - Settings</span></span><br><span data-ttu-id="b6e33-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-553">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-554">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-555">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-556">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-558">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="b6e33-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="b6e33-559">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e33-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e33-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-564">- BindingEvents</span></span><br><span data-ttu-id="b6e33-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-565">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-567">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-568">
         - File</span></span><br><span data-ttu-id="b6e33-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-570">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-571">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-574">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-575">
         - Selection</span></span><br><span data-ttu-id="b6e33-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-576">
         - Settings</span></span><br><span data-ttu-id="b6e33-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-577">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-578">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-579">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-580">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-582">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-583">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-583">- TaskPane</span></span><br><span data-ttu-id="b6e33-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e33-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e33-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-589">- BindingEvents</span></span><br><span data-ttu-id="b6e33-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-590">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-592">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-593">
         - File</span></span><br><span data-ttu-id="b6e33-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-595">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-596">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-599">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-600">
         - Selection</span></span><br><span data-ttu-id="b6e33-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-601">
         - Settings</span></span><br><span data-ttu-id="b6e33-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-602">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-603">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-604">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-605">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-607">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-608">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-608">- TaskPane</span></span><br><span data-ttu-id="b6e33-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e33-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e33-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e33-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e33-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-614">- BindingEvents</span></span><br><span data-ttu-id="b6e33-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-615">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-617">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-618">
         - File</span></span><br><span data-ttu-id="b6e33-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-620">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-621">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-624">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-625">
         - Selection</span></span><br><span data-ttu-id="b6e33-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-626">
         - Settings</span></span><br><span data-ttu-id="b6e33-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-627">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-628">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-629">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-630">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-632">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-633">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e33-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b6e33-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="b6e33-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-636">- BindingEvents</span></span><br><span data-ttu-id="b6e33-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-637">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e33-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="b6e33-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-639">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-640">
         - File</span></span><br><span data-ttu-id="b6e33-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-642">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-643">
         - MatrixBindings</span></span><br><span data-ttu-id="b6e33-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="b6e33-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b6e33-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-646">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-647">
         - Selection</span></span><br><span data-ttu-id="b6e33-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-648">
         - Settings</span></span><br><span data-ttu-id="b6e33-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-649">
         - TableBindings</span></span><br><span data-ttu-id="b6e33-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-650">
         - TableCoercion</span></span><br><span data-ttu-id="b6e33-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e33-651">
         - TextBindings</span></span><br><span data-ttu-id="b6e33-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-652">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b6e33-654">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b6e33-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b6e33-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b6e33-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e33-656">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b6e33-656">Platform</span></span></th>
    <th><span data-ttu-id="b6e33-657">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-657">Extension points</span></span></th>
    <th><span data-ttu-id="b6e33-658">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e33-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e33-660">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e33-661">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-661">- Content</span></span><br><span data-ttu-id="b6e33-662">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-662">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-665">- ActiveView</span></span><br><span data-ttu-id="b6e33-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-666">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-667">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-668">
         - File</span></span><br><span data-ttu-id="b6e33-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-669">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-670">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-671">
         - Selection</span></span><br><span data-ttu-id="b6e33-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-672">
         - Settings</span></span><br><span data-ttu-id="b6e33-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-674">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-675">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-675">- Content</span></span><br><span data-ttu-id="b6e33-676">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-676">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-679">- ActiveView</span></span><br><span data-ttu-id="b6e33-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-680">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-681">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-682">
         - File</span></span><br><span data-ttu-id="b6e33-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-683">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-684">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-685">
         - Selection</span></span><br><span data-ttu-id="b6e33-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-686">
         - Settings</span></span><br><span data-ttu-id="b6e33-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-688">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-689">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-689">- Content</span></span><br><span data-ttu-id="b6e33-690">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-690">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-693">- ActiveView</span></span><br><span data-ttu-id="b6e33-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-694">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-695">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-696">
         - File</span></span><br><span data-ttu-id="b6e33-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-697">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-698">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-699">
         - Selection</span></span><br><span data-ttu-id="b6e33-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-700">
         - Settings</span></span><br><span data-ttu-id="b6e33-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-702">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-703">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-703">- Content</span></span><br><span data-ttu-id="b6e33-704">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b6e33-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b6e33-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-706">- ActiveView</span></span><br><span data-ttu-id="b6e33-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-707">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-708">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-709">
         - File</span></span><br><span data-ttu-id="b6e33-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-710">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-711">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-712">
         - Selection</span></span><br><span data-ttu-id="b6e33-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-713">
         - Settings</span></span><br><span data-ttu-id="b6e33-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-715">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-716">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-716">- Content</span></span><br><span data-ttu-id="b6e33-717">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b6e33-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b6e33-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b6e33-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-719">- ActiveView</span></span><br><span data-ttu-id="b6e33-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-720">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-721">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-722">
         - File</span></span><br><span data-ttu-id="b6e33-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-723">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-724">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-725">
         - Selection</span></span><br><span data-ttu-id="b6e33-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-726">
         - Settings</span></span><br><span data-ttu-id="b6e33-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-728">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="b6e33-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="b6e33-729">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-729">- Content</span></span><br><span data-ttu-id="b6e33-730">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="b6e33-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-732">- ActiveView</span></span><br><span data-ttu-id="b6e33-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-733">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-734">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-735">
         - File</span></span><br><span data-ttu-id="b6e33-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-736">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-737">
         - Selection</span></span><br><span data-ttu-id="b6e33-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-738">
         - Settings</span></span><br><span data-ttu-id="b6e33-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-739">
         - TextCoercion</span></span><br><span data-ttu-id="b6e33-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-741">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-742">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-742">- Content</span></span><br><span data-ttu-id="b6e33-743">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-743">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-746">- ActiveView</span></span><br><span data-ttu-id="b6e33-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-747">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-748">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-749">
         - File</span></span><br><span data-ttu-id="b6e33-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-750">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-751">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-752">
         - Selection</span></span><br><span data-ttu-id="b6e33-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-753">
         - Settings</span></span><br><span data-ttu-id="b6e33-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-755">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-756">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-756">- Content</span></span><br><span data-ttu-id="b6e33-757">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-757">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-760">- ActiveView</span></span><br><span data-ttu-id="b6e33-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-761">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-762">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-763">
         - File</span></span><br><span data-ttu-id="b6e33-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-764">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-765">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-766">
         - Selection</span></span><br><span data-ttu-id="b6e33-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-767">
         - Settings</span></span><br><span data-ttu-id="b6e33-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-769">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="b6e33-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e33-770">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-770">- Content</span></span><br><span data-ttu-id="b6e33-771">Volet Office 
         -/td></span><span class="sxs-lookup"><span data-stu-id="b6e33-771">
         - TaskPane/td></span></span> <td> <span data-ttu-id="b6e33-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b6e33-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b6e33-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e33-773">- ActiveView</span></span><br><span data-ttu-id="b6e33-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-774">
         - CompressedFile</span></span><br><span data-ttu-id="b6e33-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-775">
         - DocumentEvents</span></span><br><span data-ttu-id="b6e33-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e33-776">
         - File</span></span><br><span data-ttu-id="b6e33-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-777">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e33-778">
         - PdfFile</span></span><br><span data-ttu-id="b6e33-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-779">
         - Selection</span></span><br><span data-ttu-id="b6e33-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-780">
         - Settings</span></span><br><span data-ttu-id="b6e33-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b6e33-782">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b6e33-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b6e33-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="b6e33-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e33-784">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b6e33-784">Platform</span></span></th>
    <th><span data-ttu-id="b6e33-785">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-785">Extension points</span></span></th>
    <th><span data-ttu-id="b6e33-786">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e33-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e33-788">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e33-789">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b6e33-789">- Content</span></span><br><span data-ttu-id="b6e33-790">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-790">
         - TaskPane</span></span><br><span data-ttu-id="b6e33-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e33-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b6e33-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e33-794">- DocumentEvents</span></span><br><span data-ttu-id="b6e33-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="b6e33-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-796">
         - ImageCoercion</span></span><br><span data-ttu-id="b6e33-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e33-797">
         - Settings</span></span><br><span data-ttu-id="b6e33-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b6e33-799">Projet</span><span class="sxs-lookup"><span data-stu-id="b6e33-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e33-800">Plateforme</span><span class="sxs-lookup"><span data-stu-id="b6e33-800">Platform</span></span></th>
    <th><span data-ttu-id="b6e33-801">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b6e33-801">Extension points</span></span></th>
    <th><span data-ttu-id="b6e33-802">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b6e33-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e33-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e33-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-804">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-805">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-807">- Selection</span></span><br><span data-ttu-id="b6e33-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-809">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-810">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-812">- Selection</span></span><br><span data-ttu-id="b6e33-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e33-814">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="b6e33-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e33-815">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b6e33-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e33-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e33-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="b6e33-817">- Selection</span></span><br><span data-ttu-id="b6e33-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e33-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b6e33-819">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b6e33-819">See also</span></span>

- [<span data-ttu-id="b6e33-820">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b6e33-821">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="b6e33-821">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b6e33-822">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="b6e33-822">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b6e33-823">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="b6e33-823">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
