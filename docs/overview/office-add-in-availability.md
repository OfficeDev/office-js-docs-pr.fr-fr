---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 28d63866a03bcae99829d3a6b6c6198059a92bdc
ms.sourcegitcommit: 4d9f3e177b0bcd62804d5045f52b03e441af244f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2019
ms.locfileid: "37440149"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d52f1-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="d52f1-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d52f1-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="d52f1-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d52f1-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="d52f1-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="d52f1-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="d52f1-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="d52f1-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d52f1-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d52f1-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d52f1-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d52f1-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d52f1-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d52f1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="d52f1-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-114">- TaskPane</span></span><br><span data-ttu-id="d52f1-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-115">
        - Content</span></span><br><span data-ttu-id="d52f1-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-116">
        - Custom Functions</span></span><br><span data-ttu-id="d52f1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="d52f1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d52f1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d52f1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-128">
        - BindingEvents</span></span><br><span data-ttu-id="d52f1-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-129">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-130">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-131">
        - File</span></span><br><span data-ttu-id="d52f1-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-132">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-134">
        - Selection</span></span><br><span data-ttu-id="d52f1-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-135">
        - Settings</span></span><br><span data-ttu-id="d52f1-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-136">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-137">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-138">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-140">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-140">Office on Windows</span></span><br><span data-ttu-id="d52f1-141">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-142">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-142">- TaskPane</span></span><br><span data-ttu-id="d52f1-143">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-143">
        - Content</span></span><br><span data-ttu-id="d52f1-144">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-144">
        - Custom Functions</span></span><br><span data-ttu-id="d52f1-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="d52f1-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d52f1-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d52f1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d52f1-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-158">
        - BindingEvents</span></span><br><span data-ttu-id="d52f1-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-159">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-160">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-161">
        - File</span></span><br><span data-ttu-id="d52f1-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-162">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-164">
        - Selection</span></span><br><span data-ttu-id="d52f1-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-165">
        - Settings</span></span><br><span data-ttu-id="d52f1-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-166">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-167">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-168">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-170">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-170">Office 2019 on Windows</span></span><br><span data-ttu-id="d52f1-171">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d52f1-172">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-172">- TaskPane</span></span><br><span data-ttu-id="d52f1-173">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-173">
        - Content</span></span><br><span data-ttu-id="d52f1-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d52f1-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-185">- BindingEvents</span></span><br><span data-ttu-id="d52f1-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-186">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-187">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-188">
        - File</span></span><br><span data-ttu-id="d52f1-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-189">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-191">
        - Selection</span></span><br><span data-ttu-id="d52f1-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-192">
        - Settings</span></span><br><span data-ttu-id="d52f1-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-193">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-194">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-195">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-197">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-197">Office 2016 on Windows</span></span><br><span data-ttu-id="d52f1-198">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d52f1-199">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-199">- TaskPane</span></span><br><span data-ttu-id="d52f1-200">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-200">
        - Content</span></span></td>
    <td><span data-ttu-id="d52f1-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d52f1-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-204">- BindingEvents</span></span><br><span data-ttu-id="d52f1-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-205">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-206">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-207">
        - File</span></span><br><span data-ttu-id="d52f1-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-208">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-210">
        - Selection</span></span><br><span data-ttu-id="d52f1-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-211">
        - Settings</span></span><br><span data-ttu-id="d52f1-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-212">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-213">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-214">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-216">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-216">Office 2013 on Windows</span></span><br><span data-ttu-id="d52f1-217">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d52f1-218">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-218">
        - TaskPane</span></span><br><span data-ttu-id="d52f1-219">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d52f1-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d52f1-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d52f1-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-222">
        - BindingEvents</span></span><br><span data-ttu-id="d52f1-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-223">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-224">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-225">
        - File</span></span><br><span data-ttu-id="d52f1-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-226">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-228">
        - Selection</span></span><br><span data-ttu-id="d52f1-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-229">
        - Settings</span></span><br><span data-ttu-id="d52f1-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-230">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-231">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-232">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-234">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d52f1-234">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d52f1-235">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d52f1-236">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-236">- TaskPane</span></span><br><span data-ttu-id="d52f1-237">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-237">
        - Content</span></span></td>
    <td><span data-ttu-id="d52f1-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d52f1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-249">- BindingEvents</span></span><br><span data-ttu-id="d52f1-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-250">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-251">
        - File</span></span><br><span data-ttu-id="d52f1-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-252">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-254">
        - Selection</span></span><br><span data-ttu-id="d52f1-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-255">
        - Settings</span></span><br><span data-ttu-id="d52f1-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-256">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-257">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-258">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-260">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-260">Office apps on Mac</span></span><br><span data-ttu-id="d52f1-261">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d52f1-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-262">- TaskPane</span></span><br><span data-ttu-id="d52f1-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-263">
        - Content</span></span><br><span data-ttu-id="d52f1-264">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-264">
        - Custom Functions</span></span><br><span data-ttu-id="d52f1-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d52f1-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d52f1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d52f1-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-278">- BindingEvents</span></span><br><span data-ttu-id="d52f1-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-279">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-280">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-281">
        - File</span></span><br><span data-ttu-id="d52f1-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-282">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-284">
        - PdfFile</span></span><br><span data-ttu-id="d52f1-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-285">
        - Selection</span></span><br><span data-ttu-id="d52f1-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-286">
        - Settings</span></span><br><span data-ttu-id="d52f1-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-287">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-288">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-289">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-291">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-291">Office 2019 for Mac</span></span><br><span data-ttu-id="d52f1-292">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d52f1-293">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-293">- TaskPane</span></span><br><span data-ttu-id="d52f1-294">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-294">
        - Content</span></span><br><span data-ttu-id="d52f1-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d52f1-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d52f1-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d52f1-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d52f1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d52f1-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d52f1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d52f1-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d52f1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-306">- BindingEvents</span></span><br><span data-ttu-id="d52f1-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-307">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-308">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-309">
        - File</span></span><br><span data-ttu-id="d52f1-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-310">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-312">
        - PdfFile</span></span><br><span data-ttu-id="d52f1-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-313">
        - Selection</span></span><br><span data-ttu-id="d52f1-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-314">
        - Settings</span></span><br><span data-ttu-id="d52f1-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-315">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-316">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-317">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-319">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="d52f1-320">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d52f1-321">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-321">- TaskPane</span></span><br><span data-ttu-id="d52f1-322">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-322">
        - Content</span></span></td>
    <td><span data-ttu-id="d52f1-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d52f1-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d52f1-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d52f1-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-326">- BindingEvents</span></span><br><span data-ttu-id="d52f1-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-327">
        - CompressedFile</span></span><br><span data-ttu-id="d52f1-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-328">
        - DocumentEvents</span></span><br><span data-ttu-id="d52f1-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-329">
        - File</span></span><br><span data-ttu-id="d52f1-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-330">
        - MatrixBindings</span></span><br><span data-ttu-id="d52f1-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-332">
        - PdfFile</span></span><br><span data-ttu-id="d52f1-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-333">
        - Selection</span></span><br><span data-ttu-id="d52f1-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-334">
        - Settings</span></span><br><span data-ttu-id="d52f1-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-335">
        - TableBindings</span></span><br><span data-ttu-id="d52f1-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-336">
        - TableCoercion</span></span><br><span data-ttu-id="d52f1-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-337">
        - TextBindings</span></span><br><span data-ttu-id="d52f1-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d52f1-339">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d52f1-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="d52f1-340">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d52f1-341">Plateforme</span><span class="sxs-lookup"><span data-stu-id="d52f1-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d52f1-342">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d52f1-343">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d52f1-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-345">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-345">Office on the web</span></span></td>
    <td><span data-ttu-id="d52f1-346">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d52f1-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-348">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-348">Office on Windows</span></span><br><span data-ttu-id="d52f1-349">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d52f1-350">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d52f1-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-352">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-352">Office for Mac</span></span><br><span data-ttu-id="d52f1-353">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="d52f1-354">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d52f1-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d52f1-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="d52f1-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="d52f1-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d52f1-357">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d52f1-357">Platform</span></span></th>
    <th><span data-ttu-id="d52f1-358">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-358">Extension points</span></span></th>
    <th><span data-ttu-id="d52f1-359">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="d52f1-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-361">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-361">Office on the web</span></span><br><span data-ttu-id="d52f1-362">(moderne)</span><span class="sxs-lookup"><span data-stu-id="d52f1-362">Modern</span></span></td>
    <td> <span data-ttu-id="d52f1-363">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-363">- Mail Read</span></span><br><span data-ttu-id="d52f1-364">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-364">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d52f1-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d52f1-373">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-374">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-374">Office on the web</span></span><br><span data-ttu-id="d52f1-375">(classique)</span><span class="sxs-lookup"><span data-stu-id="d52f1-375">Classic.</span></span></td>
    <td> <span data-ttu-id="d52f1-376">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-376">- Mail Read</span></span><br><span data-ttu-id="d52f1-377">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-377">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d52f1-385">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-386">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-386">Office on Windows</span></span><br><span data-ttu-id="d52f1-387">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-388">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-388">- Mail Read</span></span><br><span data-ttu-id="d52f1-389">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-389">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d52f1-391">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="d52f1-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d52f1-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d52f1-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d52f1-399">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-400">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-400">Office 2019 on Windows</span></span><br><span data-ttu-id="d52f1-401">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-402">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-402">- Mail Read</span></span><br><span data-ttu-id="d52f1-403">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-403">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d52f1-405">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="d52f1-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d52f1-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d52f1-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d52f1-413">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-414">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-414">Office 2016 on Windows</span></span><br><span data-ttu-id="d52f1-415">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-416">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-416">- Mail Read</span></span><br><span data-ttu-id="d52f1-417">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-417">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d52f1-419">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="d52f1-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d52f1-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d52f1-424">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-425">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-425">Office 2013 on Windows</span></span><br><span data-ttu-id="d52f1-426">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-427">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-427">- Mail Read</span></span><br><span data-ttu-id="d52f1-428">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="d52f1-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="d52f1-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d52f1-433">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-434">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="d52f1-434">Office apps on iOS</span></span><br><span data-ttu-id="d52f1-435">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-436">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-436">- Mail Read</span></span><br><span data-ttu-id="d52f1-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d52f1-443">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-444">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-444">Office apps on Mac</span></span><br><span data-ttu-id="d52f1-445">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-446">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-446">- Mail Read</span></span><br><span data-ttu-id="d52f1-447">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-447">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d52f1-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d52f1-456">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-457">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-457">Office 2019 for Mac</span></span><br><span data-ttu-id="d52f1-458">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-459">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-459">- Mail Read</span></span><br><span data-ttu-id="d52f1-460">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-460">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d52f1-468">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-469">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="d52f1-470">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-471">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-471">- Mail Read</span></span><br><span data-ttu-id="d52f1-472">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-472">
      - Mail Compose</span></span><br><span data-ttu-id="d52f1-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d52f1-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d52f1-480">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-481">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="d52f1-481">Office apps on Android</span></span><br><span data-ttu-id="d52f1-482">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-483">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="d52f1-483">- Mail Read</span></span><br><span data-ttu-id="d52f1-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d52f1-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d52f1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d52f1-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d52f1-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d52f1-490">Non disponible</span><span class="sxs-lookup"><span data-stu-id="d52f1-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="d52f1-491">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d52f1-491">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d52f1-492">La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d52f1-492">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="d52f1-493">Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="d52f1-493">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="d52f1-494">Word</span><span class="sxs-lookup"><span data-stu-id="d52f1-494">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d52f1-495">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d52f1-495">Platform</span></span></th>
    <th><span data-ttu-id="d52f1-496">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-496">Extension points</span></span></th>
    <th><span data-ttu-id="d52f1-497">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="d52f1-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-499">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-499">Office on the web</span></span></td>
    <td> <span data-ttu-id="d52f1-500">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-500">- TaskPane</span></span><br><span data-ttu-id="d52f1-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d52f1-508">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-508">- BindingEvents</span></span><br><span data-ttu-id="d52f1-509">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-509">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-510">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-510">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-511">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-511">
         - File</span></span><br><span data-ttu-id="d52f1-512">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-512">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-513">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-513">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-514">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-514">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-515">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-515">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-516">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-516">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-517">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-517">
         - Selection</span></span><br><span data-ttu-id="d52f1-518">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-518">
         - Settings</span></span><br><span data-ttu-id="d52f1-519">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-519">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-520">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-520">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-521">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-521">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-522">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-522">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-523">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-523">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-524">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-524">Office on Windows</span></span><br><span data-ttu-id="d52f1-525">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-525">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-526">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-526">- TaskPane</span></span><br><span data-ttu-id="d52f1-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d52f1-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-534">- BindingEvents</span></span><br><span data-ttu-id="d52f1-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-535">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-537">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-538">
         - File</span></span><br><span data-ttu-id="d52f1-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-540">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-543">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-544">
         - Selection</span></span><br><span data-ttu-id="d52f1-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-545">
         - Settings</span></span><br><span data-ttu-id="d52f1-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-546">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-547">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-548">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-549">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-550">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-551">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-551">Office 2019 on Windows</span></span><br><span data-ttu-id="d52f1-552">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-552">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-553">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="d52f1-553">- TaskPane</span></span><br><span data-ttu-id="d52f1-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-560">- BindingEvents</span></span><br><span data-ttu-id="d52f1-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-561">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-563">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-564">
         - File</span></span><br><span data-ttu-id="d52f1-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-566">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-569">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-570">
         - Selection</span></span><br><span data-ttu-id="d52f1-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-571">
         - Settings</span></span><br><span data-ttu-id="d52f1-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-572">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-573">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-574">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-575">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-577">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-577">Office 2016 on Windows</span></span><br><span data-ttu-id="d52f1-578">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-579">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-579">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d52f1-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-583">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-583">- BindingEvents</span></span><br><span data-ttu-id="d52f1-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-584">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-585">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-585">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-586">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-586">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-587">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-587">
         - File</span></span><br><span data-ttu-id="d52f1-588">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-588">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-589">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-589">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-590">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-590">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-591">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-591">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-592">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-592">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-593">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-593">
         - Selection</span></span><br><span data-ttu-id="d52f1-594">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-594">
         - Settings</span></span><br><span data-ttu-id="d52f1-595">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-595">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-596">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-596">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-597">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-597">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-598">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-599">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-599">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-600">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-600">Office 2013 on Windows</span></span><br><span data-ttu-id="d52f1-601">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-601">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-602">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-602">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d52f1-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d52f1-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-605">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-605">- BindingEvents</span></span><br><span data-ttu-id="d52f1-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-606">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-607">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-607">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-608">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-609">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-609">
         - File</span></span><br><span data-ttu-id="d52f1-610">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-610">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-611">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-611">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-612">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-612">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-613">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-613">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-614">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-615">
         - Selection</span></span><br><span data-ttu-id="d52f1-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-616">
         - Settings</span></span><br><span data-ttu-id="d52f1-617">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-617">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-618">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-618">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-619">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-619">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-620">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-620">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-621">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-621">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-622">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d52f1-622">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d52f1-623">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-623">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-624">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-624">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d52f1-630">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-630">- BindingEvents</span></span><br><span data-ttu-id="d52f1-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-631">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-632">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-632">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-633">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-634">
         - File</span></span><br><span data-ttu-id="d52f1-635">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-635">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-636">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-636">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-637">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-637">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-638">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-638">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-639">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-639">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-640">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-640">
         - Selection</span></span><br><span data-ttu-id="d52f1-641">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-641">
         - Settings</span></span><br><span data-ttu-id="d52f1-642">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-642">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-643">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-643">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-644">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-644">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-645">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-645">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-646">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-646">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-647">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-647">Office apps on Mac</span></span><br><span data-ttu-id="d52f1-648">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-648">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-649">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-649">- TaskPane</span></span><br><span data-ttu-id="d52f1-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="d52f1-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-657">- BindingEvents</span></span><br><span data-ttu-id="d52f1-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-658">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-660">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-661">
         - File</span></span><br><span data-ttu-id="d52f1-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-663">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-666">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-667">
         - Selection</span></span><br><span data-ttu-id="d52f1-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-668">
         - Settings</span></span><br><span data-ttu-id="d52f1-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-669">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-670">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-671">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-672">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-674">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-674">Office 2019 for Mac</span></span><br><span data-ttu-id="d52f1-675">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-675">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-676">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="d52f1-676">- TaskPane</span></span><br><span data-ttu-id="d52f1-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d52f1-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d52f1-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d52f1-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-683">- BindingEvents</span></span><br><span data-ttu-id="d52f1-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-684">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-686">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-687">
         - File</span></span><br><span data-ttu-id="d52f1-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-689">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-692">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-693">
         - Selection</span></span><br><span data-ttu-id="d52f1-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-694">
         - Settings</span></span><br><span data-ttu-id="d52f1-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-695">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-696">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-697">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-698">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-700">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-700">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="d52f1-701">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-702">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-702">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d52f1-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d52f1-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d52f1-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-706">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-706">- BindingEvents</span></span><br><span data-ttu-id="d52f1-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-707">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-708">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d52f1-708">
         - CustomXmlParts</span></span><br><span data-ttu-id="d52f1-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-709">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-710">
         - File</span></span><br><span data-ttu-id="d52f1-711">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-711">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-712">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-712">
         - MatrixBindings</span></span><br><span data-ttu-id="d52f1-713">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-713">
         - MatrixCoercion</span></span><br><span data-ttu-id="d52f1-714">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-714">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d52f1-715">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-715">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-716">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-716">
         - Selection</span></span><br><span data-ttu-id="d52f1-717">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-717">
         - Settings</span></span><br><span data-ttu-id="d52f1-718">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-718">
         - TableBindings</span></span><br><span data-ttu-id="d52f1-719">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-719">
         - TableCoercion</span></span><br><span data-ttu-id="d52f1-720">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d52f1-720">
         - TextBindings</span></span><br><span data-ttu-id="d52f1-721">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-721">
         - TextCoercion</span></span><br><span data-ttu-id="d52f1-722">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-722">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d52f1-723">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d52f1-723">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d52f1-724">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d52f1-724">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d52f1-725">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d52f1-725">Platform</span></span></th>
    <th><span data-ttu-id="d52f1-726">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-726">Extension points</span></span></th>
    <th><span data-ttu-id="d52f1-727">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-727">API requirement sets</span></span></th>
    <th><span data-ttu-id="d52f1-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-729">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-729">Office on the web</span></span></td>
    <td> <span data-ttu-id="d52f1-730">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-730">- Content</span></span><br><span data-ttu-id="d52f1-731">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-731">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d52f1-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d52f1-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-737">- ActiveView</span></span><br><span data-ttu-id="d52f1-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-738">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-739">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-740">
         - File</span></span><br><span data-ttu-id="d52f1-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-741">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-742">
         - Selection</span></span><br><span data-ttu-id="d52f1-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-743">
         - Settings</span></span><br><span data-ttu-id="d52f1-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-745">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-745">Office on Windows</span></span><br><span data-ttu-id="d52f1-746">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-747">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-747">- Content</span></span><br><span data-ttu-id="d52f1-748">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-748">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d52f1-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d52f1-754">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-754">- ActiveView</span></span><br><span data-ttu-id="d52f1-755">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-755">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-756">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-756">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-757">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-757">
         - File</span></span><br><span data-ttu-id="d52f1-758">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-758">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-759">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-759">
         - Selection</span></span><br><span data-ttu-id="d52f1-760">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-760">
         - Settings</span></span><br><span data-ttu-id="d52f1-761">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-761">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-762">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-762">Office 2019 on Windows</span></span><br><span data-ttu-id="d52f1-763">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-763">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-764">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-764">- Content</span></span><br><span data-ttu-id="d52f1-765">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-765">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-769">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-769">- ActiveView</span></span><br><span data-ttu-id="d52f1-770">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-770">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-771">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-771">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-772">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-772">
         - File</span></span><br><span data-ttu-id="d52f1-773">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-773">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-774">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-774">
         - Selection</span></span><br><span data-ttu-id="d52f1-775">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-775">
         - Settings</span></span><br><span data-ttu-id="d52f1-776">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-776">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-777">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-777">Office 2016 on Windows</span></span><br><span data-ttu-id="d52f1-778">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-778">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-779">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-779">- Content</span></span><br><span data-ttu-id="d52f1-780">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-780">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d52f1-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d52f1-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-783">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-783">- ActiveView</span></span><br><span data-ttu-id="d52f1-784">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-784">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-785">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-785">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-786">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-786">
         - File</span></span><br><span data-ttu-id="d52f1-787">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-787">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-788">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-788">
         - Selection</span></span><br><span data-ttu-id="d52f1-789">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-789">
         - Settings</span></span><br><span data-ttu-id="d52f1-790">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-790">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-791">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-791">Office 2013 on Windows</span></span><br><span data-ttu-id="d52f1-792">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-792">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-793">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-793">- Content</span></span><br><span data-ttu-id="d52f1-794">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-794">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d52f1-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d52f1-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d52f1-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-797">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-797">- ActiveView</span></span><br><span data-ttu-id="d52f1-798">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-798">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-799">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-799">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-800">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-800">
         - File</span></span><br><span data-ttu-id="d52f1-801">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-801">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-802">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-802">
         - Selection</span></span><br><span data-ttu-id="d52f1-803">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-803">
         - Settings</span></span><br><span data-ttu-id="d52f1-804">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-804">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-805">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="d52f1-805">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d52f1-806">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-806">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-807">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-807">- Content</span></span><br><span data-ttu-id="d52f1-808">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-808">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d52f1-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-812">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-812">- ActiveView</span></span><br><span data-ttu-id="d52f1-813">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-813">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-814">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-814">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-815">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-815">
         - File</span></span><br><span data-ttu-id="d52f1-816">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-816">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-817">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-817">
         - Selection</span></span><br><span data-ttu-id="d52f1-818">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-818">
         - Settings</span></span><br><span data-ttu-id="d52f1-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-819">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-820">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-820">Office apps on Mac</span></span><br><span data-ttu-id="d52f1-821">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52f1-821">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d52f1-822">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-822">- Content</span></span><br><span data-ttu-id="d52f1-823">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-823">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d52f1-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d52f1-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d52f1-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-829">- ActiveView</span></span><br><span data-ttu-id="d52f1-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-830">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-831">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-832">
         - File</span></span><br><span data-ttu-id="d52f1-833">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-833">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-834">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-834">
         - Selection</span></span><br><span data-ttu-id="d52f1-835">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-835">
         - Settings</span></span><br><span data-ttu-id="d52f1-836">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-836">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-837">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-837">Office 2019 for Mac</span></span><br><span data-ttu-id="d52f1-838">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-838">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-839">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-839">- Content</span></span><br><span data-ttu-id="d52f1-840">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-840">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-844">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-844">- ActiveView</span></span><br><span data-ttu-id="d52f1-845">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-845">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-846">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-846">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-847">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-847">
         - File</span></span><br><span data-ttu-id="d52f1-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-848">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-849">
         - Selection</span></span><br><span data-ttu-id="d52f1-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-850">
         - Settings</span></span><br><span data-ttu-id="d52f1-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-851">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-852">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-852">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="d52f1-853">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-853">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-854">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-854">- Content</span></span><br><span data-ttu-id="d52f1-855">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-855">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d52f1-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d52f1-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-858">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d52f1-858">- ActiveView</span></span><br><span data-ttu-id="d52f1-859">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-859">
         - CompressedFile</span></span><br><span data-ttu-id="d52f1-860">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-860">
         - DocumentEvents</span></span><br><span data-ttu-id="d52f1-861">
         - File</span><span class="sxs-lookup"><span data-stu-id="d52f1-861">
         - File</span></span><br><span data-ttu-id="d52f1-862">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d52f1-862">
         - PdfFile</span></span><br><span data-ttu-id="d52f1-863">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-863">
         - Selection</span></span><br><span data-ttu-id="d52f1-864">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-864">
         - Settings</span></span><br><span data-ttu-id="d52f1-865">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-865">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d52f1-866">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="d52f1-866">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d52f1-867">OneNote</span><span class="sxs-lookup"><span data-stu-id="d52f1-867">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d52f1-868">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="d52f1-868">Platform</span></span></th>
    <th><span data-ttu-id="d52f1-869">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-869">Extension points</span></span></th>
    <th><span data-ttu-id="d52f1-870">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-870">API requirement sets</span></span></th>
    <th><span data-ttu-id="d52f1-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-872">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="d52f1-872">Office on the web</span></span></td>
    <td> <span data-ttu-id="d52f1-873">- Contenu</span><span class="sxs-lookup"><span data-stu-id="d52f1-873">- Content</span></span><br><span data-ttu-id="d52f1-874">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-874">
         - TaskPane</span></span><br><span data-ttu-id="d52f1-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d52f1-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d52f1-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d52f1-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-879">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d52f1-879">- DocumentEvents</span></span><br><span data-ttu-id="d52f1-880">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-880">
         - HtmlCoercion</span></span><br><span data-ttu-id="d52f1-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d52f1-881">
         - Settings</span></span><br><span data-ttu-id="d52f1-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d52f1-883">Projet</span><span class="sxs-lookup"><span data-stu-id="d52f1-883">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d52f1-884">Plateforme</span><span class="sxs-lookup"><span data-stu-id="d52f1-884">Platform</span></span></th>
    <th><span data-ttu-id="d52f1-885">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="d52f1-885">Extension points</span></span></th>
    <th><span data-ttu-id="d52f1-886">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="d52f1-886">API requirement sets</span></span></th>
    <th><span data-ttu-id="d52f1-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="d52f1-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-888">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-888">Office 2019 on Windows</span></span><br><span data-ttu-id="d52f1-889">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-889">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-890">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-890">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-892">- Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-892">- Selection</span></span><br><span data-ttu-id="d52f1-893">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-893">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-894">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-894">Office 2016 on Windows</span></span><br><span data-ttu-id="d52f1-895">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-895">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-896">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-896">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-898">- Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-898">- Selection</span></span><br><span data-ttu-id="d52f1-899">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-899">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d52f1-900">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="d52f1-900">Office 2013 on Windows</span></span><br><span data-ttu-id="d52f1-901">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="d52f1-901">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d52f1-902">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-902">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d52f1-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d52f1-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d52f1-904">- Selection</span><span class="sxs-lookup"><span data-stu-id="d52f1-904">- Selection</span></span><br><span data-ttu-id="d52f1-905">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d52f1-905">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d52f1-906">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d52f1-906">See also</span></span>

- [<span data-ttu-id="d52f1-907">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-907">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d52f1-908">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d52f1-908">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d52f1-909">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="d52f1-909">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="d52f1-910">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="d52f1-910">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="d52f1-911">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="d52f1-911">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="d52f1-912">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="d52f1-912">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="d52f1-913">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="d52f1-913">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="d52f1-914">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="d52f1-914">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="d52f1-915">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d52f1-915">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="d52f1-916">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d52f1-916">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="d52f1-917">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="d52f1-917">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
