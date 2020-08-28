---
title: Élément OfficeTab dans le fichier manifest
description: L’élément OfficeTab définit l’onglet du ruban où votre commande de complément s’affiche.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 25e8044d8b3264bf9ee64c54487566bf11f0065e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292299"
---
# <a name="officetab-element"></a><span data-ttu-id="ac453-103">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ac453-103">OfficeTab element</span></span>

<span data-ttu-id="ac453-104">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche.</span><span class="sxs-lookup"><span data-stu-id="ac453-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="ac453-105">Il peut s’agir de l’onglet par défaut ( **domicile**, **message**ou **réunion**) ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="ac453-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="ac453-106">Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="ac453-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ac453-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="ac453-107">Child elements</span></span>

|  <span data-ttu-id="ac453-108">Élément</span><span class="sxs-lookup"><span data-stu-id="ac453-108">Element</span></span> |  <span data-ttu-id="ac453-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="ac453-109">Required</span></span>  |  <span data-ttu-id="ac453-110">Description</span><span class="sxs-lookup"><span data-stu-id="ac453-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ac453-111">Groupe</span><span class="sxs-lookup"><span data-stu-id="ac453-111">Group</span></span>      | <span data-ttu-id="ac453-112">Oui</span><span class="sxs-lookup"><span data-stu-id="ac453-112">Yes</span></span> |  <span data-ttu-id="ac453-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="ac453-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="ac453-115">Les valeurs d’onglet valides sont les suivantes `id` par application.</span><span class="sxs-lookup"><span data-stu-id="ac453-115">The following are valid tab `id` values by application.</span></span> <span data-ttu-id="ac453-116">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).</span><span class="sxs-lookup"><span data-stu-id="ac453-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="ac453-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="ac453-117">Outlook</span></span>

- <span data-ttu-id="ac453-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="ac453-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="ac453-119">Word</span><span class="sxs-lookup"><span data-stu-id="ac453-119">Word</span></span>

- <span data-ttu-id="ac453-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="ac453-120">**TabHome**</span></span>
- <span data-ttu-id="ac453-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="ac453-121">**TabInsert**</span></span>
- <span data-ttu-id="ac453-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="ac453-122">TabWordDesign</span></span>
- <span data-ttu-id="ac453-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="ac453-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="ac453-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="ac453-124">TabReferences</span></span>
- <span data-ttu-id="ac453-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="ac453-125">TabMailings</span></span>
- <span data-ttu-id="ac453-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="ac453-126">TabReviewWord</span></span>
- <span data-ttu-id="ac453-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="ac453-127">**TabView**</span></span>
- <span data-ttu-id="ac453-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="ac453-128">TabDeveloper</span></span>
- <span data-ttu-id="ac453-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="ac453-129">TabAddIns</span></span>
- <span data-ttu-id="ac453-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="ac453-130">TabBlogPost</span></span>
- <span data-ttu-id="ac453-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="ac453-131">TabBlogInsert</span></span>
- <span data-ttu-id="ac453-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="ac453-132">TabPrintPreview</span></span>
- <span data-ttu-id="ac453-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="ac453-133">TabOutlining</span></span>
- <span data-ttu-id="ac453-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="ac453-134">TabConflicts</span></span>
- <span data-ttu-id="ac453-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="ac453-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="ac453-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="ac453-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="ac453-137">Excel</span><span class="sxs-lookup"><span data-stu-id="ac453-137">Excel</span></span>

- <span data-ttu-id="ac453-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="ac453-138">**TabHome**</span></span>
- <span data-ttu-id="ac453-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="ac453-139">**TabInsert**</span></span>
- <span data-ttu-id="ac453-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="ac453-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="ac453-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="ac453-141">TabFormulas</span></span>
- <span data-ttu-id="ac453-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="ac453-142">**TabData**</span></span>
- <span data-ttu-id="ac453-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="ac453-143">**TabReview**</span></span>
- <span data-ttu-id="ac453-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="ac453-144">**TabView**</span></span>
- <span data-ttu-id="ac453-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="ac453-145">TabDeveloper</span></span>
- <span data-ttu-id="ac453-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="ac453-146">TabAddIns</span></span>
- <span data-ttu-id="ac453-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="ac453-147">TabPrintPreview</span></span>
- <span data-ttu-id="ac453-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="ac453-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="ac453-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ac453-149">PowerPoint</span></span>

- <span data-ttu-id="ac453-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="ac453-150">**TabHome**</span></span>
- <span data-ttu-id="ac453-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="ac453-151">**TabInsert**</span></span>
- <span data-ttu-id="ac453-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="ac453-152">**TabDesign**</span></span>
- <span data-ttu-id="ac453-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="ac453-153">**TabTransitions**</span></span>
- <span data-ttu-id="ac453-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="ac453-154">**TabAnimations**</span></span>
- <span data-ttu-id="ac453-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="ac453-155">TabSlideShow</span></span>
- <span data-ttu-id="ac453-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="ac453-156">TabReview</span></span>
- <span data-ttu-id="ac453-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="ac453-157">**TabView**</span></span>
- <span data-ttu-id="ac453-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="ac453-158">TabDeveloper</span></span>
- <span data-ttu-id="ac453-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="ac453-159">TabAddIns</span></span>
- <span data-ttu-id="ac453-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="ac453-160">TabPrintPreview</span></span>
- <span data-ttu-id="ac453-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="ac453-161">TabMerge</span></span>
- <span data-ttu-id="ac453-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="ac453-162">TabGrayscale</span></span>
- <span data-ttu-id="ac453-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="ac453-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="ac453-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="ac453-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="ac453-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="ac453-165">TabSlideMaster</span></span>
- <span data-ttu-id="ac453-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="ac453-166">TabHandoutMaster</span></span>
- <span data-ttu-id="ac453-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="ac453-167">TabNotesMaster</span></span>
- <span data-ttu-id="ac453-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="ac453-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="ac453-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="ac453-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="ac453-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="ac453-170">OneNote</span></span>

- <span data-ttu-id="ac453-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="ac453-171">**TabHome**</span></span>
- <span data-ttu-id="ac453-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="ac453-172">**TabInsert**</span></span>
- <span data-ttu-id="ac453-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="ac453-173">**TabView**</span></span>
- <span data-ttu-id="ac453-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="ac453-174">TabDeveloper</span></span>
- <span data-ttu-id="ac453-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="ac453-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="ac453-176">Group</span><span class="sxs-lookup"><span data-stu-id="ac453-176">Group</span></span>

<span data-ttu-id="ac453-177">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut avoir jusqu’à six contrôles.</span><span class="sxs-lookup"><span data-stu-id="ac453-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="ac453-178">L’attribut **ID** est obligatoire et chaque **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="ac453-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="ac453-179">L' **ID** est une chaîne avec un maximum de 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="ac453-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="ac453-180">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="ac453-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="ac453-181">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ac453-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
