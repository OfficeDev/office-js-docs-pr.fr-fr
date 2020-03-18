---
title: Élément OfficeTab dans le fichier manifest
description: L’élément OfficeTab définit l’onglet du ruban où votre commande de complément s’affiche.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 1d1810f3d3a206f72bf9544814a3fdaaa556476e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720490"
---
# <a name="officetab-element"></a><span data-ttu-id="c05ee-103">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="c05ee-103">OfficeTab element</span></span>

<span data-ttu-id="c05ee-104">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche.</span><span class="sxs-lookup"><span data-stu-id="c05ee-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="c05ee-105">Il peut s’agir de l’onglet par défaut ( **domicile**, **message**ou **réunion**) ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="c05ee-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="c05ee-106">Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="c05ee-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c05ee-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c05ee-107">Child elements</span></span>

|  <span data-ttu-id="c05ee-108">Élément</span><span class="sxs-lookup"><span data-stu-id="c05ee-108">Element</span></span> |  <span data-ttu-id="c05ee-109">Requis</span><span class="sxs-lookup"><span data-stu-id="c05ee-109">Required</span></span>  |  <span data-ttu-id="c05ee-110">Description</span><span class="sxs-lookup"><span data-stu-id="c05ee-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c05ee-111">Groupe</span><span class="sxs-lookup"><span data-stu-id="c05ee-111">Group</span></span>      | <span data-ttu-id="c05ee-112">Oui</span><span class="sxs-lookup"><span data-stu-id="c05ee-112">Yes</span></span> |  <span data-ttu-id="c05ee-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="c05ee-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="c05ee-115">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="c05ee-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="c05ee-116">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).</span><span class="sxs-lookup"><span data-stu-id="c05ee-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="c05ee-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="c05ee-117">Outlook</span></span>

- <span data-ttu-id="c05ee-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="c05ee-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="c05ee-119">Word</span><span class="sxs-lookup"><span data-stu-id="c05ee-119">Word</span></span>

- <span data-ttu-id="c05ee-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c05ee-120">**TabHome**</span></span>
- <span data-ttu-id="c05ee-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c05ee-121">**TabInsert**</span></span>
- <span data-ttu-id="c05ee-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="c05ee-122">TabWordDesign</span></span>
- <span data-ttu-id="c05ee-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="c05ee-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="c05ee-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="c05ee-124">TabReferences</span></span>
- <span data-ttu-id="c05ee-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="c05ee-125">TabMailings</span></span>
- <span data-ttu-id="c05ee-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="c05ee-126">TabReviewWord</span></span>
- <span data-ttu-id="c05ee-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c05ee-127">**TabView**</span></span>
- <span data-ttu-id="c05ee-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c05ee-128">TabDeveloper</span></span>
- <span data-ttu-id="c05ee-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c05ee-129">TabAddIns</span></span>
- <span data-ttu-id="c05ee-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="c05ee-130">TabBlogPost</span></span>
- <span data-ttu-id="c05ee-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="c05ee-131">TabBlogInsert</span></span>
- <span data-ttu-id="c05ee-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c05ee-132">TabPrintPreview</span></span>
- <span data-ttu-id="c05ee-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="c05ee-133">TabOutlining</span></span>
- <span data-ttu-id="c05ee-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="c05ee-134">TabConflicts</span></span>
- <span data-ttu-id="c05ee-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c05ee-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="c05ee-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="c05ee-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="c05ee-137">Excel</span><span class="sxs-lookup"><span data-stu-id="c05ee-137">Excel</span></span>

- <span data-ttu-id="c05ee-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c05ee-138">**TabHome**</span></span>
- <span data-ttu-id="c05ee-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c05ee-139">**TabInsert**</span></span>
- <span data-ttu-id="c05ee-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="c05ee-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="c05ee-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="c05ee-141">TabFormulas</span></span>
- <span data-ttu-id="c05ee-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="c05ee-142">**TabData**</span></span>
- <span data-ttu-id="c05ee-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="c05ee-143">**TabReview**</span></span>
- <span data-ttu-id="c05ee-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c05ee-144">**TabView**</span></span>
- <span data-ttu-id="c05ee-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c05ee-145">TabDeveloper</span></span>
- <span data-ttu-id="c05ee-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c05ee-146">TabAddIns</span></span>
- <span data-ttu-id="c05ee-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c05ee-147">TabPrintPreview</span></span>
- <span data-ttu-id="c05ee-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c05ee-148">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="c05ee-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c05ee-149">PowerPoint</span></span>

- <span data-ttu-id="c05ee-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c05ee-150">**TabHome**</span></span>
- <span data-ttu-id="c05ee-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c05ee-151">**TabInsert**</span></span>
- <span data-ttu-id="c05ee-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="c05ee-152">**TabDesign**</span></span>
- <span data-ttu-id="c05ee-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="c05ee-153">**TabTransitions**</span></span>
- <span data-ttu-id="c05ee-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="c05ee-154">**TabAnimations**</span></span>
- <span data-ttu-id="c05ee-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="c05ee-155">TabSlideShow</span></span>
- <span data-ttu-id="c05ee-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="c05ee-156">TabReview</span></span>
- <span data-ttu-id="c05ee-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c05ee-157">**TabView**</span></span>
- <span data-ttu-id="c05ee-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c05ee-158">TabDeveloper</span></span>
- <span data-ttu-id="c05ee-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c05ee-159">TabAddIns</span></span>
- <span data-ttu-id="c05ee-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c05ee-160">TabPrintPreview</span></span>
- <span data-ttu-id="c05ee-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="c05ee-161">TabMerge</span></span>
- <span data-ttu-id="c05ee-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="c05ee-162">TabGrayscale</span></span>
- <span data-ttu-id="c05ee-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="c05ee-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="c05ee-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="c05ee-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="c05ee-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="c05ee-165">TabSlideMaster</span></span>
- <span data-ttu-id="c05ee-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="c05ee-166">TabHandoutMaster</span></span>
- <span data-ttu-id="c05ee-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="c05ee-167">TabNotesMaster</span></span>
- <span data-ttu-id="c05ee-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c05ee-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="c05ee-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="c05ee-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="c05ee-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="c05ee-170">OneNote</span></span>

- <span data-ttu-id="c05ee-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c05ee-171">**TabHome**</span></span>
- <span data-ttu-id="c05ee-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c05ee-172">**TabInsert**</span></span>
- <span data-ttu-id="c05ee-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c05ee-173">**TabView**</span></span>
- <span data-ttu-id="c05ee-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c05ee-174">TabDeveloper</span></span>
- <span data-ttu-id="c05ee-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c05ee-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="c05ee-176">Group</span><span class="sxs-lookup"><span data-stu-id="c05ee-176">Group</span></span>

<span data-ttu-id="c05ee-177">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut avoir jusqu’à six contrôles.</span><span class="sxs-lookup"><span data-stu-id="c05ee-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="c05ee-178">L’attribut **ID** est obligatoire et chaque **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="c05ee-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="c05ee-179">L' **ID** est une chaîne avec un maximum de 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="c05ee-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="c05ee-180">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="c05ee-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="c05ee-181">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="c05ee-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
