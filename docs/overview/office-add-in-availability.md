---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: d88f7c1b9daa201d9b6bc5cfa69ac3125bf127b1
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2019
ms.locfileid: "35630535"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b493a-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="b493a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b493a-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="b493a-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b493a-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="b493a-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b493a-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b493a-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b493a-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b493a-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b493a-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b493a-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b493a-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b493a-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b493a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b493a-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-114">- TaskPane</span></span><br><span data-ttu-id="b493a-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-115">
        - Content</span></span><br><span data-ttu-id="b493a-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-116">
        - Custom Functions</span></span><br><span data-ttu-id="b493a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b493a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b493a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b493a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b493a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b493a-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-130">
        - BindingEvents</span></span><br><span data-ttu-id="b493a-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-131">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-132">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-133">
        - File</span></span><br><span data-ttu-id="b493a-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-134">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-136">
        - Selection</span></span><br><span data-ttu-id="b493a-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-137">
        - Settings</span></span><br><span data-ttu-id="b493a-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-138">
        - TableBindings</span></span><br><span data-ttu-id="b493a-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-139">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-140">
        - TextBindings</span></span><br><span data-ttu-id="b493a-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-142">Office on Windows</span></span><br><span data-ttu-id="b493a-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-144">- TaskPane</span></span><br><span data-ttu-id="b493a-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-145">
        - Content</span></span><br><span data-ttu-id="b493a-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-146">
        - Custom Functions</span></span><br><span data-ttu-id="b493a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="b493a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b493a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b493a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b493a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b493a-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-160">
        - BindingEvents</span></span><br><span data-ttu-id="b493a-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-161">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-162">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-163">
        - File</span></span><br><span data-ttu-id="b493a-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-164">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-166">
        - Selection</span></span><br><span data-ttu-id="b493a-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-167">
        - Settings</span></span><br><span data-ttu-id="b493a-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-168">
        - TableBindings</span></span><br><span data-ttu-id="b493a-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-169">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-170">
        - TextBindings</span></span><br><span data-ttu-id="b493a-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-172">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-172">Office 2019 on Windows</span></span><br><span data-ttu-id="b493a-173">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b493a-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-174">- TaskPane</span></span><br><span data-ttu-id="b493a-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-175">
        - Content</span></span><br><span data-ttu-id="b493a-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b493a-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-187">- BindingEvents</span></span><br><span data-ttu-id="b493a-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-188">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-189">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-190">
        - File</span></span><br><span data-ttu-id="b493a-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-191">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-193">
        - Selection</span></span><br><span data-ttu-id="b493a-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-194">
        - Settings</span></span><br><span data-ttu-id="b493a-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-195">
        - TableBindings</span></span><br><span data-ttu-id="b493a-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-196">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-197">
        - TextBindings</span></span><br><span data-ttu-id="b493a-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-199">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-199">Office 2016 on Windows</span></span><br><span data-ttu-id="b493a-200">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b493a-201">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-201">- TaskPane</span></span><br><span data-ttu-id="b493a-202">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-202">
        - Content</span></span></td>
    <td><span data-ttu-id="b493a-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b493a-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-206">- BindingEvents</span></span><br><span data-ttu-id="b493a-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-207">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-208">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-209">
        - File</span></span><br><span data-ttu-id="b493a-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-210">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-212">
        - Selection</span></span><br><span data-ttu-id="b493a-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-213">
        - Settings</span></span><br><span data-ttu-id="b493a-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-214">
        - TableBindings</span></span><br><span data-ttu-id="b493a-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-215">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-216">
        - TextBindings</span></span><br><span data-ttu-id="b493a-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-218">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-218">Office 2013 on Windows</span></span><br><span data-ttu-id="b493a-219">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b493a-220">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-220">
        - TaskPane</span></span><br><span data-ttu-id="b493a-221">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b493a-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b493a-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b493a-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-224">
        - BindingEvents</span></span><br><span data-ttu-id="b493a-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-225">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-226">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-227">
        - File</span></span><br><span data-ttu-id="b493a-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-228">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-230">
        - Selection</span></span><br><span data-ttu-id="b493a-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-231">
        - Settings</span></span><br><span data-ttu-id="b493a-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-232">
        - TableBindings</span></span><br><span data-ttu-id="b493a-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-233">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-234">
        - TextBindings</span></span><br><span data-ttu-id="b493a-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-236">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b493a-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="b493a-237">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b493a-238">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-238">- TaskPane</span></span><br><span data-ttu-id="b493a-239">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-239">
        - Content</span></span><br><span data-ttu-id="b493a-240">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b493a-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b493a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b493a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-252">- BindingEvents</span></span><br><span data-ttu-id="b493a-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-253">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-254">
        - File</span></span><br><span data-ttu-id="b493a-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-255">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-257">
        - Selection</span></span><br><span data-ttu-id="b493a-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-258">
        - Settings</span></span><br><span data-ttu-id="b493a-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-259">
        - TableBindings</span></span><br><span data-ttu-id="b493a-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-260">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-261">
        - TextBindings</span></span><br><span data-ttu-id="b493a-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-263">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-263">Office apps on Mac</span></span><br><span data-ttu-id="b493a-264">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b493a-265">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-265">- TaskPane</span></span><br><span data-ttu-id="b493a-266">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-266">
        - Content</span></span><br><span data-ttu-id="b493a-267">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-267">
        - Custom Functions</span></span><br><span data-ttu-id="b493a-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b493a-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b493a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b493a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b493a-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-281">- BindingEvents</span></span><br><span data-ttu-id="b493a-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-282">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-283">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-284">
        - File</span></span><br><span data-ttu-id="b493a-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-285">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-287">
        - PdfFile</span></span><br><span data-ttu-id="b493a-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-288">
        - Selection</span></span><br><span data-ttu-id="b493a-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-289">
        - Settings</span></span><br><span data-ttu-id="b493a-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-290">
        - TableBindings</span></span><br><span data-ttu-id="b493a-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-291">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-292">
        - TextBindings</span></span><br><span data-ttu-id="b493a-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-294">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-294">Office 2019 for Mac</span></span><br><span data-ttu-id="b493a-295">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b493a-296">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-296">- TaskPane</span></span><br><span data-ttu-id="b493a-297">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-297">
        - Content</span></span><br><span data-ttu-id="b493a-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b493a-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b493a-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b493a-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b493a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b493a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b493a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b493a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b493a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b493a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-309">- BindingEvents</span></span><br><span data-ttu-id="b493a-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-310">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-311">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-312">
        - File</span></span><br><span data-ttu-id="b493a-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-313">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-315">
        - PdfFile</span></span><br><span data-ttu-id="b493a-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-316">
        - Selection</span></span><br><span data-ttu-id="b493a-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-317">
        - Settings</span></span><br><span data-ttu-id="b493a-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-318">
        - TableBindings</span></span><br><span data-ttu-id="b493a-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-319">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-320">
        - TextBindings</span></span><br><span data-ttu-id="b493a-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-322">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="b493a-323">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b493a-324">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-324">- TaskPane</span></span><br><span data-ttu-id="b493a-325">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-325">
        - Content</span></span></td>
    <td><span data-ttu-id="b493a-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b493a-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b493a-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b493a-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-329">- BindingEvents</span></span><br><span data-ttu-id="b493a-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-330">
        - CompressedFile</span></span><br><span data-ttu-id="b493a-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-331">
        - DocumentEvents</span></span><br><span data-ttu-id="b493a-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="b493a-332">
        - File</span></span><br><span data-ttu-id="b493a-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-333">
        - MatrixBindings</span></span><br><span data-ttu-id="b493a-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="b493a-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-335">
        - PdfFile</span></span><br><span data-ttu-id="b493a-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-336">
        - Selection</span></span><br><span data-ttu-id="b493a-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-337">
        - Settings</span></span><br><span data-ttu-id="b493a-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-338">
        - TableBindings</span></span><br><span data-ttu-id="b493a-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-339">
        - TableCoercion</span></span><br><span data-ttu-id="b493a-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-340">
        - TextBindings</span></span><br><span data-ttu-id="b493a-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b493a-342">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b493a-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="b493a-343">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b493a-344">Plateforme</span><span class="sxs-lookup"><span data-stu-id="b493a-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b493a-345">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b493a-346">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b493a-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-348">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-348">Office on the web</span></span></td>
    <td><span data-ttu-id="b493a-349">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b493a-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-351">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-351">Office on Windows</span></span><br><span data-ttu-id="b493a-352">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b493a-353">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b493a-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-355">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-355">Office for Mac</span></span><br><span data-ttu-id="b493a-356">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b493a-357">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b493a-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b493a-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b493a-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="b493a-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b493a-360">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b493a-360">Platform</span></span></th>
    <th><span data-ttu-id="b493a-361">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-361">Extension points</span></span></th>
    <th><span data-ttu-id="b493a-362">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="b493a-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-364">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-364">Office on the web</span></span><br><span data-ttu-id="b493a-365">(nouveau)</span><span class="sxs-lookup"><span data-stu-id="b493a-365">New</span></span></td>
    <td> <span data-ttu-id="b493a-366">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-366">- Mail Read</span></span><br><span data-ttu-id="b493a-367">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-367">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b493a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b493a-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-377">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-377">Office on the web</span></span><br><span data-ttu-id="b493a-378">(classique)</span><span class="sxs-lookup"><span data-stu-id="b493a-378">Classic.</span></span></td>
    <td> <span data-ttu-id="b493a-379">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-379">- Mail Read</span></span><br><span data-ttu-id="b493a-380">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-380">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b493a-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-389">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-389">Office on Windows</span></span><br><span data-ttu-id="b493a-390">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-391">- Mail Read</span></span><br><span data-ttu-id="b493a-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-392">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b493a-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b493a-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b493a-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b493a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b493a-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-403">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-403">Office 2019 on Windows</span></span><br><span data-ttu-id="b493a-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-405">- Mail Read</span></span><br><span data-ttu-id="b493a-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-406">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b493a-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b493a-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b493a-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b493a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b493a-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-417">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-417">Office 2016 on Windows</span></span><br><span data-ttu-id="b493a-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-419">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-419">- Mail Read</span></span><br><span data-ttu-id="b493a-420">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-420">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b493a-422">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="b493a-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b493a-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b493a-427">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-428">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-428">Office 2013 on Windows</span></span><br><span data-ttu-id="b493a-429">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-430">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-430">- Mail Read</span></span><br><span data-ttu-id="b493a-431">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="b493a-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b493a-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b493a-436">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-437">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="b493a-437">Office apps on iOS</span></span><br><span data-ttu-id="b493a-438">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-439">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-439">- Mail Read</span></span><br><span data-ttu-id="b493a-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b493a-446">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-447">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-447">Office apps on Mac</span></span><br><span data-ttu-id="b493a-448">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-449">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-449">- Mail Read</span></span><br><span data-ttu-id="b493a-450">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-450">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b493a-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b493a-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b493a-459">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-460">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-460">Office 2019 for Mac</span></span><br><span data-ttu-id="b493a-461">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-462">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-462">- Mail Read</span></span><br><span data-ttu-id="b493a-463">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-463">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b493a-471">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-472">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="b493a-473">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-474">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-474">- Mail Read</span></span><br><span data-ttu-id="b493a-475">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="b493a-475">
      - Mail Compose</span></span><br><span data-ttu-id="b493a-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b493a-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b493a-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b493a-483">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-484">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="b493a-484">Office apps on Android</span></span><br><span data-ttu-id="b493a-485">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-486">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="b493a-486">- Mail Read</span></span><br><span data-ttu-id="b493a-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b493a-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b493a-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b493a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b493a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b493a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b493a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b493a-493">Non disponible</span><span class="sxs-lookup"><span data-stu-id="b493a-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b493a-494">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b493a-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b493a-495">Word</span><span class="sxs-lookup"><span data-stu-id="b493a-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b493a-496">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b493a-496">Platform</span></span></th>
    <th><span data-ttu-id="b493a-497">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-497">Extension points</span></span></th>
    <th><span data-ttu-id="b493a-498">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="b493a-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-500">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="b493a-501">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-501">- TaskPane</span></span><br><span data-ttu-id="b493a-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b493a-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-509">- BindingEvents</span></span><br><span data-ttu-id="b493a-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-511">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-512">
         - File</span></span><br><span data-ttu-id="b493a-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-514">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-517">
         - PdfFile</span></span><br><span data-ttu-id="b493a-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-518">
         - Selection</span></span><br><span data-ttu-id="b493a-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-519">
         - Settings</span></span><br><span data-ttu-id="b493a-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-520">
         - TableBindings</span></span><br><span data-ttu-id="b493a-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-521">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-522">
         - TextBindings</span></span><br><span data-ttu-id="b493a-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-523">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-525">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-525">Office on Windows</span></span><br><span data-ttu-id="b493a-526">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-527">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-527">- TaskPane</span></span><br><span data-ttu-id="b493a-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b493a-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-535">- BindingEvents</span></span><br><span data-ttu-id="b493a-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-536">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-538">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-539">
         - File</span></span><br><span data-ttu-id="b493a-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-541">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-544">
         - PdfFile</span></span><br><span data-ttu-id="b493a-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-545">
         - Selection</span></span><br><span data-ttu-id="b493a-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-546">
         - Settings</span></span><br><span data-ttu-id="b493a-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-547">
         - TableBindings</span></span><br><span data-ttu-id="b493a-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-548">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-549">
         - TextBindings</span></span><br><span data-ttu-id="b493a-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-550">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-552">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-552">Office 2019 on Windows</span></span><br><span data-ttu-id="b493a-553">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-554">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="b493a-554">- TaskPane</span></span><br><span data-ttu-id="b493a-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-561">- BindingEvents</span></span><br><span data-ttu-id="b493a-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-562">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-564">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-565">
         - File</span></span><br><span data-ttu-id="b493a-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-567">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-570">
         - PdfFile</span></span><br><span data-ttu-id="b493a-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-571">
         - Selection</span></span><br><span data-ttu-id="b493a-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-572">
         - Settings</span></span><br><span data-ttu-id="b493a-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-573">
         - TableBindings</span></span><br><span data-ttu-id="b493a-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-574">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-575">
         - TextBindings</span></span><br><span data-ttu-id="b493a-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-576">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-578">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-578">Office 2016 on Windows</span></span><br><span data-ttu-id="b493a-579">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-580">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b493a-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-584">- BindingEvents</span></span><br><span data-ttu-id="b493a-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-585">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-587">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-588">
         - File</span></span><br><span data-ttu-id="b493a-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-590">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-593">
         - PdfFile</span></span><br><span data-ttu-id="b493a-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-594">
         - Selection</span></span><br><span data-ttu-id="b493a-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-595">
         - Settings</span></span><br><span data-ttu-id="b493a-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-596">
         - TableBindings</span></span><br><span data-ttu-id="b493a-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-597">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-598">
         - TextBindings</span></span><br><span data-ttu-id="b493a-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-599">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-601">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-601">Office 2013 on Windows</span></span><br><span data-ttu-id="b493a-602">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-603">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b493a-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b493a-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-606">- BindingEvents</span></span><br><span data-ttu-id="b493a-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-607">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-609">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-610">
         - File</span></span><br><span data-ttu-id="b493a-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-612">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-615">
         - PdfFile</span></span><br><span data-ttu-id="b493a-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-616">
         - Selection</span></span><br><span data-ttu-id="b493a-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-617">
         - Settings</span></span><br><span data-ttu-id="b493a-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-618">
         - TableBindings</span></span><br><span data-ttu-id="b493a-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-619">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-620">
         - TextBindings</span></span><br><span data-ttu-id="b493a-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-621">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-623">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b493a-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="b493a-624">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-625">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b493a-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-631">- BindingEvents</span></span><br><span data-ttu-id="b493a-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-632">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-634">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-635">
         - File</span></span><br><span data-ttu-id="b493a-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-637">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-640">
         - PdfFile</span></span><br><span data-ttu-id="b493a-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-641">
         - Selection</span></span><br><span data-ttu-id="b493a-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-642">
         - Settings</span></span><br><span data-ttu-id="b493a-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-643">
         - TableBindings</span></span><br><span data-ttu-id="b493a-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-644">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-645">
         - TextBindings</span></span><br><span data-ttu-id="b493a-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-646">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-648">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-648">Office apps on Mac</span></span><br><span data-ttu-id="b493a-649">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-650">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-650">- TaskPane</span></span><br><span data-ttu-id="b493a-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b493a-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-658">- BindingEvents</span></span><br><span data-ttu-id="b493a-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-659">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-661">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-662">
         - File</span></span><br><span data-ttu-id="b493a-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-664">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-667">
         - PdfFile</span></span><br><span data-ttu-id="b493a-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-668">
         - Selection</span></span><br><span data-ttu-id="b493a-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-669">
         - Settings</span></span><br><span data-ttu-id="b493a-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-670">
         - TableBindings</span></span><br><span data-ttu-id="b493a-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-671">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-672">
         - TextBindings</span></span><br><span data-ttu-id="b493a-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-673">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-675">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-675">Office 2019 for Mac</span></span><br><span data-ttu-id="b493a-676">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-677">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="b493a-677">- TaskPane</span></span><br><span data-ttu-id="b493a-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b493a-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b493a-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b493a-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b493a-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-684">- BindingEvents</span></span><br><span data-ttu-id="b493a-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-685">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-687">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-688">
         - File</span></span><br><span data-ttu-id="b493a-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-690">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-693">
         - PdfFile</span></span><br><span data-ttu-id="b493a-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-694">
         - Selection</span></span><br><span data-ttu-id="b493a-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-695">
         - Settings</span></span><br><span data-ttu-id="b493a-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-696">
         - TableBindings</span></span><br><span data-ttu-id="b493a-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-697">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-698">
         - TextBindings</span></span><br><span data-ttu-id="b493a-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-699">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-701">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="b493a-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b493a-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b493a-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b493a-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-707">- BindingEvents</span></span><br><span data-ttu-id="b493a-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-708">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b493a-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="b493a-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-710">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-711">
         - File</span></span><br><span data-ttu-id="b493a-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-713">
         - MatrixBindings</span></span><br><span data-ttu-id="b493a-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="b493a-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b493a-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-716">
         - PdfFile</span></span><br><span data-ttu-id="b493a-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-717">
         - Selection</span></span><br><span data-ttu-id="b493a-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-718">
         - Settings</span></span><br><span data-ttu-id="b493a-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-719">
         - TableBindings</span></span><br><span data-ttu-id="b493a-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-720">
         - TableCoercion</span></span><br><span data-ttu-id="b493a-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b493a-721">
         - TextBindings</span></span><br><span data-ttu-id="b493a-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-722">
         - TextCoercion</span></span><br><span data-ttu-id="b493a-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b493a-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b493a-724">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b493a-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b493a-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b493a-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b493a-726">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b493a-726">Platform</span></span></th>
    <th><span data-ttu-id="b493a-727">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-727">Extension points</span></span></th>
    <th><span data-ttu-id="b493a-728">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="b493a-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-730">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="b493a-731">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-731">- Content</span></span><br><span data-ttu-id="b493a-732">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-732">
         - TaskPane</span></span><br><span data-ttu-id="b493a-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b493a-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-737">- ActiveView</span></span><br><span data-ttu-id="b493a-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-738">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-739">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-740">
         - File</span></span><br><span data-ttu-id="b493a-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-741">
         - PdfFile</span></span><br><span data-ttu-id="b493a-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-742">
         - Selection</span></span><br><span data-ttu-id="b493a-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-743">
         - Settings</span></span><br><span data-ttu-id="b493a-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-745">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-745">Office on Windows</span></span><br><span data-ttu-id="b493a-746">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-747">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-747">- Content</span></span><br><span data-ttu-id="b493a-748">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-748">
         - TaskPane</span></span><br><span data-ttu-id="b493a-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b493a-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-753">- ActiveView</span></span><br><span data-ttu-id="b493a-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-754">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-755">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-756">
         - File</span></span><br><span data-ttu-id="b493a-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-757">
         - PdfFile</span></span><br><span data-ttu-id="b493a-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-758">
         - Selection</span></span><br><span data-ttu-id="b493a-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-759">
         - Settings</span></span><br><span data-ttu-id="b493a-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-761">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-761">Office 2019 on Windows</span></span><br><span data-ttu-id="b493a-762">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-763">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-763">- Content</span></span><br><span data-ttu-id="b493a-764">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-764">
         - TaskPane</span></span><br><span data-ttu-id="b493a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-768">- ActiveView</span></span><br><span data-ttu-id="b493a-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-769">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-770">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-771">
         - File</span></span><br><span data-ttu-id="b493a-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-772">
         - PdfFile</span></span><br><span data-ttu-id="b493a-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-773">
         - Selection</span></span><br><span data-ttu-id="b493a-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-774">
         - Settings</span></span><br><span data-ttu-id="b493a-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-776">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-776">Office 2016 on Windows</span></span><br><span data-ttu-id="b493a-777">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-778">- Content</span></span><br><span data-ttu-id="b493a-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b493a-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b493a-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-782">- ActiveView</span></span><br><span data-ttu-id="b493a-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-783">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-784">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-785">
         - File</span></span><br><span data-ttu-id="b493a-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-786">
         - PdfFile</span></span><br><span data-ttu-id="b493a-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-787">
         - Selection</span></span><br><span data-ttu-id="b493a-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-788">
         - Settings</span></span><br><span data-ttu-id="b493a-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-790">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-790">Office 2013 on Windows</span></span><br><span data-ttu-id="b493a-791">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-792">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-792">- Content</span></span><br><span data-ttu-id="b493a-793">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b493a-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b493a-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b493a-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-796">- ActiveView</span></span><br><span data-ttu-id="b493a-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-797">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-798">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-799">
         - File</span></span><br><span data-ttu-id="b493a-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-800">
         - PdfFile</span></span><br><span data-ttu-id="b493a-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-801">
         - Selection</span></span><br><span data-ttu-id="b493a-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-802">
         - Settings</span></span><br><span data-ttu-id="b493a-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-804">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b493a-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="b493a-805">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-806">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-806">- Content</span></span><br><span data-ttu-id="b493a-807">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-810">- ActiveView</span></span><br><span data-ttu-id="b493a-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-811">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-812">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-813">
         - File</span></span><br><span data-ttu-id="b493a-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-814">
         - PdfFile</span></span><br><span data-ttu-id="b493a-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-815">
         - Selection</span></span><br><span data-ttu-id="b493a-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-816">
         - Settings</span></span><br><span data-ttu-id="b493a-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-818">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-818">Office apps on Mac</span></span><br><span data-ttu-id="b493a-819">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="b493a-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b493a-820">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-820">- Content</span></span><br><span data-ttu-id="b493a-821">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-821">
         - TaskPane</span></span><br><span data-ttu-id="b493a-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b493a-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b493a-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b493a-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-826">- ActiveView</span></span><br><span data-ttu-id="b493a-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-827">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-828">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-829">
         - File</span></span><br><span data-ttu-id="b493a-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-830">
         - PdfFile</span></span><br><span data-ttu-id="b493a-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-831">
         - Selection</span></span><br><span data-ttu-id="b493a-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-832">
         - Settings</span></span><br><span data-ttu-id="b493a-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-834">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-834">Office 2019 for Mac</span></span><br><span data-ttu-id="b493a-835">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-836">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-836">- Content</span></span><br><span data-ttu-id="b493a-837">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-837">
         - TaskPane</span></span><br><span data-ttu-id="b493a-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-841">- ActiveView</span></span><br><span data-ttu-id="b493a-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-842">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-843">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-844">
         - File</span></span><br><span data-ttu-id="b493a-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-845">
         - PdfFile</span></span><br><span data-ttu-id="b493a-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-846">
         - Selection</span></span><br><span data-ttu-id="b493a-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-847">
         - Settings</span></span><br><span data-ttu-id="b493a-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-849">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="b493a-850">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-851">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-851">- Content</span></span><br><span data-ttu-id="b493a-852">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b493a-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b493a-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b493a-855">- ActiveView</span></span><br><span data-ttu-id="b493a-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b493a-856">
         - CompressedFile</span></span><br><span data-ttu-id="b493a-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-857">
         - DocumentEvents</span></span><br><span data-ttu-id="b493a-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="b493a-858">
         - File</span></span><br><span data-ttu-id="b493a-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b493a-859">
         - PdfFile</span></span><br><span data-ttu-id="b493a-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-860">
         - Selection</span></span><br><span data-ttu-id="b493a-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-861">
         - Settings</span></span><br><span data-ttu-id="b493a-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b493a-863">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="b493a-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b493a-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="b493a-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b493a-865">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="b493a-865">Platform</span></span></th>
    <th><span data-ttu-id="b493a-866">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-866">Extension points</span></span></th>
    <th><span data-ttu-id="b493a-867">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="b493a-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-869">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b493a-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="b493a-870">- Contenu</span><span class="sxs-lookup"><span data-stu-id="b493a-870">- Content</span></span><br><span data-ttu-id="b493a-871">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-871">
         - TaskPane</span></span><br><span data-ttu-id="b493a-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="b493a-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b493a-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b493a-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b493a-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b493a-876">- DocumentEvents</span></span><br><span data-ttu-id="b493a-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="b493a-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b493a-878">
         - Settings</span></span><br><span data-ttu-id="b493a-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b493a-880">Projet</span><span class="sxs-lookup"><span data-stu-id="b493a-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b493a-881">Plateforme</span><span class="sxs-lookup"><span data-stu-id="b493a-881">Platform</span></span></th>
    <th><span data-ttu-id="b493a-882">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b493a-882">Extension points</span></span></th>
    <th><span data-ttu-id="b493a-883">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="b493a-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="b493a-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="b493a-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-885">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-885">Office 2019 on Windows</span></span><br><span data-ttu-id="b493a-886">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-887">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-889">- Selection</span></span><br><span data-ttu-id="b493a-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-891">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-891">Office 2016 on Windows</span></span><br><span data-ttu-id="b493a-892">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-893">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-895">- Selection</span></span><br><span data-ttu-id="b493a-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b493a-897">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b493a-897">Office 2013 on Windows</span></span><br><span data-ttu-id="b493a-898">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="b493a-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b493a-899">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="b493a-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b493a-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b493a-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b493a-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="b493a-901">- Selection</span></span><br><span data-ttu-id="b493a-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b493a-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b493a-903">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b493a-903">See also</span></span>

- [<span data-ttu-id="b493a-904">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b493a-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b493a-905">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b493a-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="b493a-906">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="b493a-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b493a-907">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="b493a-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b493a-908">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="b493a-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="b493a-909">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="b493a-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b493a-910">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="b493a-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b493a-911">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="b493a-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b493a-912">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b493a-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b493a-913">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b493a-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b493a-914">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="b493a-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
