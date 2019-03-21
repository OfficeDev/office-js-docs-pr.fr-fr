---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: fe5b1d1278d2c14192fb6fd212f24bb08571d35d
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691124"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4d07a-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="4d07a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4d07a-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="4d07a-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4d07a-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="4d07a-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="4d07a-108">Le numéro de build d’un achat définitif d’Office 2019 est 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="4d07a-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="4d07a-109">Excel</span><span class="sxs-lookup"><span data-stu-id="4d07a-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4d07a-110">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4d07a-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4d07a-111">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4d07a-112">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4d07a-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d07a-114">Office Online</span></span></td>
    <td> <span data-ttu-id="4d07a-115">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-115">- TaskPane</span></span><br><span data-ttu-id="4d07a-116">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-116">
        - Content</span></span><br><span data-ttu-id="4d07a-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d07a-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d07a-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-127">
        - BindingEvents</span></span><br><span data-ttu-id="4d07a-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-128">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-129">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-130">
        - File</span></span><br><span data-ttu-id="4d07a-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-131">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-133">
        - Selection</span></span><br><span data-ttu-id="4d07a-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-134">
        - Settings</span></span><br><span data-ttu-id="4d07a-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-135">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-136">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-137">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-139">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-140">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-140">- TaskPane</span></span><br><span data-ttu-id="4d07a-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-141">
        - Content</span></span><br><span data-ttu-id="4d07a-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d07a-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d07a-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-152">
        - BindingEvents</span></span><br><span data-ttu-id="4d07a-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-153">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-154">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-155">
        - File</span></span><br><span data-ttu-id="4d07a-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-156">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-158">
        - Selection</span></span><br><span data-ttu-id="4d07a-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-159">
        - Settings</span></span><br><span data-ttu-id="4d07a-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-160">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-161">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-162">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-164">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="4d07a-165">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-165">- TaskPane</span></span><br><span data-ttu-id="4d07a-166">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-166">
        - Content</span></span><br><span data-ttu-id="4d07a-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d07a-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-177">- BindingEvents</span></span><br><span data-ttu-id="4d07a-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-178">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-179">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-180">
        - File</span></span><br><span data-ttu-id="4d07a-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-181">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-182">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-184">
        - Selection</span></span><br><span data-ttu-id="4d07a-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-185">
        - Settings</span></span><br><span data-ttu-id="4d07a-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-186">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-187">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-188">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-190">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4d07a-191">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-191">- TaskPane</span></span><br><span data-ttu-id="4d07a-192">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-192">
        - Content</span></span></td>
    <td><span data-ttu-id="4d07a-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4d07a-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-195">- BindingEvents</span></span><br><span data-ttu-id="4d07a-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-196">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-197">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-198">
        - File</span></span><br><span data-ttu-id="4d07a-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-199">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-200">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-202">
        - Selection</span></span><br><span data-ttu-id="4d07a-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-203">
        - Settings</span></span><br><span data-ttu-id="4d07a-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-204">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-205">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-206">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-208">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4d07a-209">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-209">
        - TaskPane</span></span><br><span data-ttu-id="4d07a-210">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4d07a-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d07a-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="4d07a-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-212">
        - BindingEvents</span></span><br><span data-ttu-id="4d07a-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-213">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-214">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-215">
        - File</span></span><br><span data-ttu-id="4d07a-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-216">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-217">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-219">
        - Selection</span></span><br><span data-ttu-id="4d07a-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-220">
        - Settings</span></span><br><span data-ttu-id="4d07a-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-221">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-222">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-223">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-225">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="4d07a-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="4d07a-226">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-226">- TaskPane</span></span><br><span data-ttu-id="4d07a-227">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-227">
        - Content</span></span></td>
    <td><span data-ttu-id="4d07a-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-237">- BindingEvents</span></span><br><span data-ttu-id="4d07a-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-238">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-239">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-240">
        - File</span></span><br><span data-ttu-id="4d07a-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-241">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-242">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-244">
        - Selection</span></span><br><span data-ttu-id="4d07a-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-245">
        - Settings</span></span><br><span data-ttu-id="4d07a-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-246">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-247">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-248">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-250">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="4d07a-251">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-251">- TaskPane</span></span><br><span data-ttu-id="4d07a-252">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-252">
        - Content</span></span><br><span data-ttu-id="4d07a-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d07a-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-263">- BindingEvents</span></span><br><span data-ttu-id="4d07a-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-264">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-265">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-266">
        - File</span></span><br><span data-ttu-id="4d07a-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-267">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-268">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-270">
        - PdfFile</span></span><br><span data-ttu-id="4d07a-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-271">
        - Selection</span></span><br><span data-ttu-id="4d07a-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-272">
        - Settings</span></span><br><span data-ttu-id="4d07a-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-273">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-274">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-275">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-277">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="4d07a-278">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-278">- TaskPane</span></span><br><span data-ttu-id="4d07a-279">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-279">
        - Content</span></span><br><span data-ttu-id="4d07a-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d07a-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d07a-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d07a-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d07a-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d07a-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d07a-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d07a-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d07a-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d07a-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-290">- BindingEvents</span></span><br><span data-ttu-id="4d07a-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-291">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-292">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-293">
        - File</span></span><br><span data-ttu-id="4d07a-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-294">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-295">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-297">
        - PdfFile</span></span><br><span data-ttu-id="4d07a-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-298">
        - Selection</span></span><br><span data-ttu-id="4d07a-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-299">
        - Settings</span></span><br><span data-ttu-id="4d07a-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-300">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-301">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-302">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-304">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4d07a-305">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-305">- TaskPane</span></span><br><span data-ttu-id="4d07a-306">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-306">
        - Content</span></span></td>
    <td><span data-ttu-id="4d07a-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d07a-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4d07a-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-309">- BindingEvents</span></span><br><span data-ttu-id="4d07a-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-310">
        - CompressedFile</span></span><br><span data-ttu-id="4d07a-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-311">
        - DocumentEvents</span></span><br><span data-ttu-id="4d07a-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-312">
        - File</span></span><br><span data-ttu-id="4d07a-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-313">
        - ImageCoercion</span></span><br><span data-ttu-id="4d07a-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-314">
        - MatrixBindings</span></span><br><span data-ttu-id="4d07a-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-316">
        - PdfFile</span></span><br><span data-ttu-id="4d07a-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-317">
        - Selection</span></span><br><span data-ttu-id="4d07a-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-318">
        - Settings</span></span><br><span data-ttu-id="4d07a-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-319">
        - TableBindings</span></span><br><span data-ttu-id="4d07a-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-320">
        - TableCoercion</span></span><br><span data-ttu-id="4d07a-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-321">
        - TextBindings</span></span><br><span data-ttu-id="4d07a-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d07a-323">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="4d07a-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="4d07a-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="4d07a-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d07a-325">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4d07a-325">Platform</span></span></th>
    <th><span data-ttu-id="4d07a-326">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-326">Extension points</span></span></th>
    <th><span data-ttu-id="4d07a-327">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d07a-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d07a-329">Office Online</span></span></td>
    <td> <span data-ttu-id="4d07a-330">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-330">- Mail Read</span></span><br><span data-ttu-id="4d07a-331">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-331">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d07a-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d07a-340">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-341">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-342">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-342">- Mail Read</span></span><br><span data-ttu-id="4d07a-343">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-343">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d07a-345">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="4d07a-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d07a-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d07a-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d07a-353">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-354">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-355">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-355">- Mail Read</span></span><br><span data-ttu-id="4d07a-356">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-356">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d07a-358">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="4d07a-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d07a-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d07a-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d07a-366">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-367">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-368">- Mail Read</span></span><br><span data-ttu-id="4d07a-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-369">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d07a-371">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="4d07a-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d07a-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4d07a-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-377">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-378">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-378">- Mail Read</span></span><br><span data-ttu-id="4d07a-379">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4d07a-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4d07a-384">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-385">Office 365 pour iOS</span><span class="sxs-lookup"><span data-stu-id="4d07a-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="4d07a-386">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-386">- Mail Read</span></span><br><span data-ttu-id="4d07a-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d07a-393">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-394">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-395">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-395">- Mail Read</span></span><br><span data-ttu-id="4d07a-396">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-396">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d07a-404">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-405">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-406">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-406">- Mail Read</span></span><br><span data-ttu-id="4d07a-407">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-407">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d07a-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-416">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-417">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-417">- Mail Read</span></span><br><span data-ttu-id="4d07a-418">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-418">
      - Mail Compose</span></span><br><span data-ttu-id="4d07a-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d07a-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d07a-426">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-427">Office 365 pour Android</span><span class="sxs-lookup"><span data-stu-id="4d07a-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="4d07a-428">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4d07a-428">- Mail Read</span></span><br><span data-ttu-id="4d07a-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d07a-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d07a-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d07a-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d07a-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d07a-435">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4d07a-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d07a-436">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="4d07a-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="4d07a-437">Word</span><span class="sxs-lookup"><span data-stu-id="4d07a-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d07a-438">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4d07a-438">Platform</span></span></th>
    <th><span data-ttu-id="4d07a-439">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-439">Extension points</span></span></th>
    <th><span data-ttu-id="4d07a-440">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d07a-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d07a-442">Office Online</span></span></td>
    <td> <span data-ttu-id="4d07a-443">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-443">- TaskPane</span></span><br><span data-ttu-id="4d07a-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-449">- BindingEvents</span></span><br><span data-ttu-id="4d07a-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-451">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-452">
         - File</span></span><br><span data-ttu-id="4d07a-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-454">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-455">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-458">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-459">
         - Selection</span></span><br><span data-ttu-id="4d07a-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-460">
         - Settings</span></span><br><span data-ttu-id="4d07a-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-461">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-462">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-463">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-464">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-466">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-467">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-467">- TaskPane</span></span><br><span data-ttu-id="4d07a-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-473">- BindingEvents</span></span><br><span data-ttu-id="4d07a-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-474">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-476">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-477">
         - File</span></span><br><span data-ttu-id="4d07a-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-479">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-480">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-483">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-484">
         - Selection</span></span><br><span data-ttu-id="4d07a-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-485">
         - Settings</span></span><br><span data-ttu-id="4d07a-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-486">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-487">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-488">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-489">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-491">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-492">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-492">- TaskPane</span></span><br><span data-ttu-id="4d07a-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-498">- BindingEvents</span></span><br><span data-ttu-id="4d07a-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-499">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-501">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-502">
         - File</span></span><br><span data-ttu-id="4d07a-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-504">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-505">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-508">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-509">
         - Selection</span></span><br><span data-ttu-id="4d07a-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-510">
         - Settings</span></span><br><span data-ttu-id="4d07a-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-511">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-512">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-513">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-514">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-516">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-517">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4d07a-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-520">- BindingEvents</span></span><br><span data-ttu-id="4d07a-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-521">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-523">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-524">
         - File</span></span><br><span data-ttu-id="4d07a-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-526">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-527">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-530">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-531">
         - Selection</span></span><br><span data-ttu-id="4d07a-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-532">
         - Settings</span></span><br><span data-ttu-id="4d07a-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-533">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-534">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-535">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-536">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-538">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-539">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d07a-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d07a-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-541">- BindingEvents</span></span><br><span data-ttu-id="4d07a-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-542">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-544">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-545">
         - File</span></span><br><span data-ttu-id="4d07a-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-547">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-548">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-551">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-552">
         - Selection</span></span><br><span data-ttu-id="4d07a-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-553">
         - Settings</span></span><br><span data-ttu-id="4d07a-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-554">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-555">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-556">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-557">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-559">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="4d07a-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4d07a-560">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d07a-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d07a-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-565">- BindingEvents</span></span><br><span data-ttu-id="4d07a-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-566">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-568">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-569">
         - File</span></span><br><span data-ttu-id="4d07a-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-571">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-572">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-575">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-576">
         - Selection</span></span><br><span data-ttu-id="4d07a-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-577">
         - Settings</span></span><br><span data-ttu-id="4d07a-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-578">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-579">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-580">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-581">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-583">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-584">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-584">- TaskPane</span></span><br><span data-ttu-id="4d07a-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d07a-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d07a-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-590">- BindingEvents</span></span><br><span data-ttu-id="4d07a-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-591">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-593">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-594">
         - File</span></span><br><span data-ttu-id="4d07a-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-596">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-597">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-600">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-601">
         - Selection</span></span><br><span data-ttu-id="4d07a-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-602">
         - Settings</span></span><br><span data-ttu-id="4d07a-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-603">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-604">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-605">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-606">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-608">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-609">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-609">- TaskPane</span></span><br><span data-ttu-id="4d07a-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d07a-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d07a-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d07a-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d07a-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-615">- BindingEvents</span></span><br><span data-ttu-id="4d07a-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-616">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-618">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-619">
         - File</span></span><br><span data-ttu-id="4d07a-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-621">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-622">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-625">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-626">
         - Selection</span></span><br><span data-ttu-id="4d07a-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-627">
         - Settings</span></span><br><span data-ttu-id="4d07a-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-628">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-629">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-630">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-631">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-633">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-634">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d07a-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d07a-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4d07a-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-637">- BindingEvents</span></span><br><span data-ttu-id="4d07a-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-638">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d07a-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d07a-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-640">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-641">
         - File</span></span><br><span data-ttu-id="4d07a-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-643">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-644">
         - MatrixBindings</span></span><br><span data-ttu-id="4d07a-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d07a-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d07a-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-647">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-648">
         - Selection</span></span><br><span data-ttu-id="4d07a-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-649">
         - Settings</span></span><br><span data-ttu-id="4d07a-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-650">
         - TableBindings</span></span><br><span data-ttu-id="4d07a-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-651">
         - TableCoercion</span></span><br><span data-ttu-id="4d07a-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d07a-652">
         - TextBindings</span></span><br><span data-ttu-id="4d07a-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-653">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4d07a-655">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="4d07a-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4d07a-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4d07a-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d07a-657">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4d07a-657">Platform</span></span></th>
    <th><span data-ttu-id="4d07a-658">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-658">Extension points</span></span></th>
    <th><span data-ttu-id="4d07a-659">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d07a-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d07a-661">Office Online</span></span></td>
    <td> <span data-ttu-id="4d07a-662">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-662">- Content</span></span><br><span data-ttu-id="4d07a-663">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-663">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-666">- ActiveView</span></span><br><span data-ttu-id="4d07a-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-667">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-668">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-669">
         - File</span></span><br><span data-ttu-id="4d07a-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-670">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-671">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-672">
         - Selection</span></span><br><span data-ttu-id="4d07a-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-673">
         - Settings</span></span><br><span data-ttu-id="4d07a-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-675">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-676">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-676">- Content</span></span><br><span data-ttu-id="4d07a-677">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-677">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-680">- ActiveView</span></span><br><span data-ttu-id="4d07a-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-681">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-682">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-683">
         - File</span></span><br><span data-ttu-id="4d07a-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-684">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-685">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-686">
         - Selection</span></span><br><span data-ttu-id="4d07a-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-687">
         - Settings</span></span><br><span data-ttu-id="4d07a-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-689">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-690">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-690">- Content</span></span><br><span data-ttu-id="4d07a-691">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-691">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-694">- ActiveView</span></span><br><span data-ttu-id="4d07a-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-695">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-696">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-697">
         - File</span></span><br><span data-ttu-id="4d07a-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-698">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-699">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-700">
         - Selection</span></span><br><span data-ttu-id="4d07a-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-701">
         - Settings</span></span><br><span data-ttu-id="4d07a-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-703">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-704">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-704">- Content</span></span><br><span data-ttu-id="4d07a-705">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d07a-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d07a-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-707">- ActiveView</span></span><br><span data-ttu-id="4d07a-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-708">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-709">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-710">
         - File</span></span><br><span data-ttu-id="4d07a-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-711">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-712">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-713">
         - Selection</span></span><br><span data-ttu-id="4d07a-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-714">
         - Settings</span></span><br><span data-ttu-id="4d07a-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-716">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-717">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-717">- Content</span></span><br><span data-ttu-id="4d07a-718">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4d07a-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d07a-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d07a-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-720">- ActiveView</span></span><br><span data-ttu-id="4d07a-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-721">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-722">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-723">
         - File</span></span><br><span data-ttu-id="4d07a-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-724">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-725">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-726">
         - Selection</span></span><br><span data-ttu-id="4d07a-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-727">
         - Settings</span></span><br><span data-ttu-id="4d07a-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-729">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="4d07a-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4d07a-730">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-730">- Content</span></span><br><span data-ttu-id="4d07a-731">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4d07a-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-733">- ActiveView</span></span><br><span data-ttu-id="4d07a-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-734">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-735">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-736">
         - File</span></span><br><span data-ttu-id="4d07a-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-737">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-738">
         - Selection</span></span><br><span data-ttu-id="4d07a-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-739">
         - Settings</span></span><br><span data-ttu-id="4d07a-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-740">
         - TextCoercion</span></span><br><span data-ttu-id="4d07a-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-742">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-743">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-743">- Content</span></span><br><span data-ttu-id="4d07a-744">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-744">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-747">- ActiveView</span></span><br><span data-ttu-id="4d07a-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-748">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-749">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-750">
         - File</span></span><br><span data-ttu-id="4d07a-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-751">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-752">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-753">
         - Selection</span></span><br><span data-ttu-id="4d07a-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-754">
         - Settings</span></span><br><span data-ttu-id="4d07a-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-756">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-757">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-757">- Content</span></span><br><span data-ttu-id="4d07a-758">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-758">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-761">- ActiveView</span></span><br><span data-ttu-id="4d07a-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-762">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-763">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-764">
         - File</span></span><br><span data-ttu-id="4d07a-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-765">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-766">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-767">
         - Selection</span></span><br><span data-ttu-id="4d07a-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-768">
         - Settings</span></span><br><span data-ttu-id="4d07a-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-770">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4d07a-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d07a-771">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-771">- Content</span></span><br><span data-ttu-id="4d07a-772">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d07a-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d07a-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d07a-774">- ActiveView</span></span><br><span data-ttu-id="4d07a-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-775">
         - CompressedFile</span></span><br><span data-ttu-id="4d07a-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-776">
         - DocumentEvents</span></span><br><span data-ttu-id="4d07a-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d07a-777">
         - File</span></span><br><span data-ttu-id="4d07a-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-778">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d07a-779">
         - PdfFile</span></span><br><span data-ttu-id="4d07a-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-780">
         - Selection</span></span><br><span data-ttu-id="4d07a-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-781">
         - Settings</span></span><br><span data-ttu-id="4d07a-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d07a-783">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="4d07a-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4d07a-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="4d07a-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d07a-785">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4d07a-785">Platform</span></span></th>
    <th><span data-ttu-id="4d07a-786">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-786">Extension points</span></span></th>
    <th><span data-ttu-id="4d07a-787">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d07a-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d07a-789">Office Online</span></span></td>
    <td> <span data-ttu-id="4d07a-790">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4d07a-790">- Content</span></span><br><span data-ttu-id="4d07a-791">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-791">
         - TaskPane</span></span><br><span data-ttu-id="4d07a-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d07a-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4d07a-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d07a-795">- DocumentEvents</span></span><br><span data-ttu-id="4d07a-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d07a-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-797">
         - ImageCoercion</span></span><br><span data-ttu-id="4d07a-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d07a-798">
         - Settings</span></span><br><span data-ttu-id="4d07a-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4d07a-800">Projet</span><span class="sxs-lookup"><span data-stu-id="4d07a-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d07a-801">Plateforme</span><span class="sxs-lookup"><span data-stu-id="4d07a-801">Platform</span></span></th>
    <th><span data-ttu-id="4d07a-802">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4d07a-802">Extension points</span></span></th>
    <th><span data-ttu-id="4d07a-803">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4d07a-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d07a-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4d07a-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-805">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-806">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-808">- Selection</span></span><br><span data-ttu-id="4d07a-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-810">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-811">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-813">- Selection</span></span><br><span data-ttu-id="4d07a-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d07a-815">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4d07a-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d07a-816">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d07a-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d07a-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d07a-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d07a-818">- Selection</span></span><br><span data-ttu-id="4d07a-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d07a-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4d07a-820">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4d07a-820">See also</span></span>

- [<span data-ttu-id="4d07a-821">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4d07a-822">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="4d07a-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4d07a-823">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="4d07a-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4d07a-824">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="4d07a-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
