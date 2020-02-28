---
title: Élément OfficeTab dans le fichier manifest
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b8458233ba93e98fe0bd8d51f5734b1fece65864
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324833"
---
# <a name="officetab-element"></a><span data-ttu-id="dde87-102">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="dde87-102">OfficeTab element</span></span>

<span data-ttu-id="dde87-103">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche.</span><span class="sxs-lookup"><span data-stu-id="dde87-103">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="dde87-104">Il peut s’agir de l’onglet par défaut ( **domicile**, **message**ou **réunion**) ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="dde87-104">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="dde87-105">Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="dde87-105">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dde87-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="dde87-106">Child elements</span></span>

|  <span data-ttu-id="dde87-107">Élément</span><span class="sxs-lookup"><span data-stu-id="dde87-107">Element</span></span> |  <span data-ttu-id="dde87-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="dde87-108">Required</span></span>  |  <span data-ttu-id="dde87-109">Description</span><span class="sxs-lookup"><span data-stu-id="dde87-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dde87-110">Groupe</span><span class="sxs-lookup"><span data-stu-id="dde87-110">Group</span></span>      | <span data-ttu-id="dde87-111">Oui</span><span class="sxs-lookup"><span data-stu-id="dde87-111">Yes</span></span> |  <span data-ttu-id="dde87-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="dde87-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="dde87-114">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="dde87-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="dde87-115">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).</span><span class="sxs-lookup"><span data-stu-id="dde87-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="dde87-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="dde87-116">Outlook</span></span>

- <span data-ttu-id="dde87-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="dde87-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="dde87-118">Word</span><span class="sxs-lookup"><span data-stu-id="dde87-118">Word</span></span>

- <span data-ttu-id="dde87-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dde87-119">**TabHome**</span></span>
- <span data-ttu-id="dde87-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dde87-120">**TabInsert**</span></span>
- <span data-ttu-id="dde87-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="dde87-121">TabWordDesign</span></span>
- <span data-ttu-id="dde87-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="dde87-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="dde87-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="dde87-123">TabReferences</span></span>
- <span data-ttu-id="dde87-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="dde87-124">TabMailings</span></span>
- <span data-ttu-id="dde87-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="dde87-125">TabReviewWord</span></span>
- <span data-ttu-id="dde87-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dde87-126">**TabView**</span></span>
- <span data-ttu-id="dde87-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dde87-127">TabDeveloper</span></span>
- <span data-ttu-id="dde87-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dde87-128">TabAddIns</span></span>
- <span data-ttu-id="dde87-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="dde87-129">TabBlogPost</span></span>
- <span data-ttu-id="dde87-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="dde87-130">TabBlogInsert</span></span>
- <span data-ttu-id="dde87-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dde87-131">TabPrintPreview</span></span>
- <span data-ttu-id="dde87-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="dde87-132">TabOutlining</span></span>
- <span data-ttu-id="dde87-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="dde87-133">TabConflicts</span></span>
- <span data-ttu-id="dde87-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dde87-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dde87-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dde87-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="dde87-136">Excel</span><span class="sxs-lookup"><span data-stu-id="dde87-136">Excel</span></span>

- <span data-ttu-id="dde87-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dde87-137">**TabHome**</span></span>
- <span data-ttu-id="dde87-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dde87-138">**TabInsert**</span></span>
- <span data-ttu-id="dde87-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="dde87-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="dde87-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="dde87-140">TabFormulas</span></span>
- <span data-ttu-id="dde87-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="dde87-141">**TabData**</span></span>
- <span data-ttu-id="dde87-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="dde87-142">**TabReview**</span></span>
- <span data-ttu-id="dde87-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dde87-143">**TabView**</span></span>
- <span data-ttu-id="dde87-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dde87-144">TabDeveloper</span></span>
- <span data-ttu-id="dde87-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dde87-145">TabAddIns</span></span>
- <span data-ttu-id="dde87-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dde87-146">TabPrintPreview</span></span>
- <span data-ttu-id="dde87-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dde87-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="dde87-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dde87-148">PowerPoint</span></span>

- <span data-ttu-id="dde87-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dde87-149">**TabHome**</span></span>
- <span data-ttu-id="dde87-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dde87-150">**TabInsert**</span></span>
- <span data-ttu-id="dde87-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="dde87-151">**TabDesign**</span></span>
- <span data-ttu-id="dde87-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="dde87-152">**TabTransitions**</span></span>
- <span data-ttu-id="dde87-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="dde87-153">**TabAnimations**</span></span>
- <span data-ttu-id="dde87-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="dde87-154">TabSlideShow</span></span>
- <span data-ttu-id="dde87-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="dde87-155">TabReview</span></span>
- <span data-ttu-id="dde87-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dde87-156">**TabView**</span></span>
- <span data-ttu-id="dde87-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dde87-157">TabDeveloper</span></span>
- <span data-ttu-id="dde87-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dde87-158">TabAddIns</span></span>
- <span data-ttu-id="dde87-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dde87-159">TabPrintPreview</span></span>
- <span data-ttu-id="dde87-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="dde87-160">TabMerge</span></span>
- <span data-ttu-id="dde87-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="dde87-161">TabGrayscale</span></span>
- <span data-ttu-id="dde87-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="dde87-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="dde87-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dde87-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="dde87-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="dde87-164">TabSlideMaster</span></span>
- <span data-ttu-id="dde87-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="dde87-165">TabHandoutMaster</span></span>
- <span data-ttu-id="dde87-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="dde87-166">TabNotesMaster</span></span>
- <span data-ttu-id="dde87-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dde87-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dde87-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="dde87-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="dde87-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="dde87-169">OneNote</span></span>

- <span data-ttu-id="dde87-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dde87-170">**TabHome**</span></span>
- <span data-ttu-id="dde87-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dde87-171">**TabInsert**</span></span>
- <span data-ttu-id="dde87-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dde87-172">**TabView**</span></span>
- <span data-ttu-id="dde87-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dde87-173">TabDeveloper</span></span>
- <span data-ttu-id="dde87-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dde87-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="dde87-175">Group</span><span class="sxs-lookup"><span data-stu-id="dde87-175">Group</span></span>

<span data-ttu-id="dde87-176">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut avoir jusqu’à six contrôles.</span><span class="sxs-lookup"><span data-stu-id="dde87-176">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="dde87-177">L’attribut **ID** est obligatoire et chaque **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="dde87-177">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="dde87-178">L' **ID** est une chaîne avec un maximum de 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="dde87-178">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="dde87-179">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="dde87-179">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="dde87-180">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="dde87-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
