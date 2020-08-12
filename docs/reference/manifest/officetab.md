---
title: Élément OfficeTab dans le fichier manifest
description: L’élément OfficeTab définit l’onglet du ruban où votre commande de complément s’affiche.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9b07ce1e57329e796545610e0c61a2c11d1ed55d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641438"
---
# <a name="officetab-element"></a><span data-ttu-id="9ac6a-103">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9ac6a-103">OfficeTab element</span></span>

<span data-ttu-id="9ac6a-104">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="9ac6a-105">Il peut s’agir de l’onglet par défaut ( **domicile**, **message**ou **réunion**) ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="9ac6a-106">Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9ac6a-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9ac6a-107">Child elements</span></span>

|  <span data-ttu-id="9ac6a-108">Élément</span><span class="sxs-lookup"><span data-stu-id="9ac6a-108">Element</span></span> |  <span data-ttu-id="9ac6a-109">Requis</span><span class="sxs-lookup"><span data-stu-id="9ac6a-109">Required</span></span>  |  <span data-ttu-id="9ac6a-110">Description</span><span class="sxs-lookup"><span data-stu-id="9ac6a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9ac6a-111">Groupe</span><span class="sxs-lookup"><span data-stu-id="9ac6a-111">Group</span></span>      | <span data-ttu-id="9ac6a-112">Oui</span><span class="sxs-lookup"><span data-stu-id="9ac6a-112">Yes</span></span> |  <span data-ttu-id="9ac6a-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="9ac6a-115">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="9ac6a-116">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).</span><span class="sxs-lookup"><span data-stu-id="9ac6a-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="9ac6a-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="9ac6a-117">Outlook</span></span>

- <span data-ttu-id="9ac6a-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="9ac6a-119">Word</span><span class="sxs-lookup"><span data-stu-id="9ac6a-119">Word</span></span>

- <span data-ttu-id="9ac6a-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-120">**TabHome**</span></span>
- <span data-ttu-id="9ac6a-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-121">**TabInsert**</span></span>
- <span data-ttu-id="9ac6a-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="9ac6a-122">TabWordDesign</span></span>
- <span data-ttu-id="9ac6a-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="9ac6a-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="9ac6a-124">TabReferences</span></span>
- <span data-ttu-id="9ac6a-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="9ac6a-125">TabMailings</span></span>
- <span data-ttu-id="9ac6a-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="9ac6a-126">TabReviewWord</span></span>
- <span data-ttu-id="9ac6a-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-127">**TabView**</span></span>
- <span data-ttu-id="9ac6a-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9ac6a-128">TabDeveloper</span></span>
- <span data-ttu-id="9ac6a-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9ac6a-129">TabAddIns</span></span>
- <span data-ttu-id="9ac6a-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="9ac6a-130">TabBlogPost</span></span>
- <span data-ttu-id="9ac6a-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="9ac6a-131">TabBlogInsert</span></span>
- <span data-ttu-id="9ac6a-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9ac6a-132">TabPrintPreview</span></span>
- <span data-ttu-id="9ac6a-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="9ac6a-133">TabOutlining</span></span>
- <span data-ttu-id="9ac6a-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="9ac6a-134">TabConflicts</span></span>
- <span data-ttu-id="9ac6a-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9ac6a-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="9ac6a-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="9ac6a-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="9ac6a-137">Excel</span><span class="sxs-lookup"><span data-stu-id="9ac6a-137">Excel</span></span>

- <span data-ttu-id="9ac6a-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-138">**TabHome**</span></span>
- <span data-ttu-id="9ac6a-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-139">**TabInsert**</span></span>
- <span data-ttu-id="9ac6a-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="9ac6a-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="9ac6a-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="9ac6a-141">TabFormulas</span></span>
- <span data-ttu-id="9ac6a-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-142">**TabData**</span></span>
- <span data-ttu-id="9ac6a-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-143">**TabReview**</span></span>
- <span data-ttu-id="9ac6a-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-144">**TabView**</span></span>
- <span data-ttu-id="9ac6a-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9ac6a-145">TabDeveloper</span></span>
- <span data-ttu-id="9ac6a-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9ac6a-146">TabAddIns</span></span>
- <span data-ttu-id="9ac6a-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9ac6a-147">TabPrintPreview</span></span>
- <span data-ttu-id="9ac6a-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9ac6a-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="9ac6a-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9ac6a-149">PowerPoint</span></span>

- <span data-ttu-id="9ac6a-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-150">**TabHome**</span></span>
- <span data-ttu-id="9ac6a-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-151">**TabInsert**</span></span>
- <span data-ttu-id="9ac6a-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-152">**TabDesign**</span></span>
- <span data-ttu-id="9ac6a-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-153">**TabTransitions**</span></span>
- <span data-ttu-id="9ac6a-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-154">**TabAnimations**</span></span>
- <span data-ttu-id="9ac6a-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="9ac6a-155">TabSlideShow</span></span>
- <span data-ttu-id="9ac6a-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="9ac6a-156">TabReview</span></span>
- <span data-ttu-id="9ac6a-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-157">**TabView**</span></span>
- <span data-ttu-id="9ac6a-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9ac6a-158">TabDeveloper</span></span>
- <span data-ttu-id="9ac6a-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9ac6a-159">TabAddIns</span></span>
- <span data-ttu-id="9ac6a-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9ac6a-160">TabPrintPreview</span></span>
- <span data-ttu-id="9ac6a-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="9ac6a-161">TabMerge</span></span>
- <span data-ttu-id="9ac6a-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="9ac6a-162">TabGrayscale</span></span>
- <span data-ttu-id="9ac6a-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="9ac6a-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="9ac6a-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="9ac6a-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="9ac6a-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="9ac6a-165">TabSlideMaster</span></span>
- <span data-ttu-id="9ac6a-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="9ac6a-166">TabHandoutMaster</span></span>
- <span data-ttu-id="9ac6a-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="9ac6a-167">TabNotesMaster</span></span>
- <span data-ttu-id="9ac6a-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9ac6a-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="9ac6a-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="9ac6a-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="9ac6a-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="9ac6a-170">OneNote</span></span>

- <span data-ttu-id="9ac6a-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-171">**TabHome**</span></span>
- <span data-ttu-id="9ac6a-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-172">**TabInsert**</span></span>
- <span data-ttu-id="9ac6a-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9ac6a-173">**TabView**</span></span>
- <span data-ttu-id="9ac6a-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9ac6a-174">TabDeveloper</span></span>
- <span data-ttu-id="9ac6a-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9ac6a-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="9ac6a-176">Group</span><span class="sxs-lookup"><span data-stu-id="9ac6a-176">Group</span></span>

<span data-ttu-id="9ac6a-177">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut avoir jusqu’à six contrôles.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="9ac6a-178">L’attribut **ID** est obligatoire et chaque **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="9ac6a-179">L' **ID** est une chaîne avec un maximum de 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="9ac6a-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="9ac6a-180">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="9ac6a-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="9ac6a-181">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9ac6a-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
