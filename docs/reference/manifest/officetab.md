---
title: Élément OfficeTab dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432871"
---
# <a name="officetab-element"></a><span data-ttu-id="2b419-102">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="2b419-102">OfficeTab element</span></span>

<span data-ttu-id="2b419-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="2b419-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="2b419-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2b419-106">Child elements</span></span>

|  <span data-ttu-id="2b419-107">Élément</span><span class="sxs-lookup"><span data-stu-id="2b419-107">Element</span></span> |  <span data-ttu-id="2b419-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2b419-108">Required</span></span>  |  <span data-ttu-id="2b419-109">Description</span><span class="sxs-lookup"><span data-stu-id="2b419-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2b419-110">Group</span><span class="sxs-lookup"><span data-stu-id="2b419-110">Group</span></span>      | <span data-ttu-id="2b419-111">Oui</span><span class="sxs-lookup"><span data-stu-id="2b419-111">Yes</span></span> |  <span data-ttu-id="2b419-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="2b419-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="2b419-114">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="2b419-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="2b419-115">Les valeurs en **gras** sont prises en charge à la fois sur le bureau et en ligne (par exemple, Word 2016 pour Windows et Word Online).</span><span class="sxs-lookup"><span data-stu-id="2b419-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="2b419-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="2b419-116">Outlook</span></span>

- <span data-ttu-id="2b419-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="2b419-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="2b419-118">Word</span><span class="sxs-lookup"><span data-stu-id="2b419-118">Word</span></span>

- <span data-ttu-id="2b419-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2b419-119">**TabHome**</span></span>
- <span data-ttu-id="2b419-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2b419-120">**TabInsert**</span></span>
- <span data-ttu-id="2b419-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="2b419-121">TabWordDesign</span></span>
- <span data-ttu-id="2b419-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="2b419-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="2b419-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="2b419-123">TabReferences</span></span>
- <span data-ttu-id="2b419-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="2b419-124">TabMailings</span></span>
- <span data-ttu-id="2b419-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="2b419-125">TabReviewWord</span></span>
- <span data-ttu-id="2b419-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2b419-126">**TabView**</span></span>
- <span data-ttu-id="2b419-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2b419-127">TabDeveloper</span></span>
- <span data-ttu-id="2b419-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2b419-128">TabAddIns</span></span>
- <span data-ttu-id="2b419-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="2b419-129">TabBlogPost</span></span>
- <span data-ttu-id="2b419-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="2b419-130">TabBlogInsert</span></span>
- <span data-ttu-id="2b419-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2b419-131">TabPrintPreview</span></span>
- <span data-ttu-id="2b419-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="2b419-132">TabOutlining</span></span>
- <span data-ttu-id="2b419-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="2b419-133">TabConflicts</span></span>
- <span data-ttu-id="2b419-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2b419-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2b419-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2b419-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="2b419-136">Excel</span><span class="sxs-lookup"><span data-stu-id="2b419-136">Excel</span></span>

- <span data-ttu-id="2b419-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2b419-137">**TabHome**</span></span>
- <span data-ttu-id="2b419-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2b419-138">**TabInsert**</span></span>
- <span data-ttu-id="2b419-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="2b419-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="2b419-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="2b419-140">TabFormulas</span></span>
- <span data-ttu-id="2b419-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="2b419-141">**TabData**</span></span>
- <span data-ttu-id="2b419-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="2b419-142">**TabReview**</span></span>
- <span data-ttu-id="2b419-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2b419-143">**TabView**</span></span>
- <span data-ttu-id="2b419-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2b419-144">TabDeveloper</span></span>
- <span data-ttu-id="2b419-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2b419-145">TabAddIns</span></span>
- <span data-ttu-id="2b419-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2b419-146">TabPrintPreview</span></span>
- <span data-ttu-id="2b419-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2b419-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="2b419-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2b419-148">PowerPoint</span></span>

- <span data-ttu-id="2b419-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2b419-149">**TabHome**</span></span>
- <span data-ttu-id="2b419-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2b419-150">**TabInsert**</span></span>
- <span data-ttu-id="2b419-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="2b419-151">**TabDesign**</span></span>
- <span data-ttu-id="2b419-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="2b419-152">**TabTransitions**</span></span>
- <span data-ttu-id="2b419-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="2b419-153">**TabAnimations**</span></span>
- <span data-ttu-id="2b419-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="2b419-154">TabSlideShow</span></span>
- <span data-ttu-id="2b419-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="2b419-155">TabReview</span></span>
- <span data-ttu-id="2b419-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2b419-156">**TabView**</span></span>
- <span data-ttu-id="2b419-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2b419-157">TabDeveloper</span></span>
- <span data-ttu-id="2b419-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2b419-158">TabAddIns</span></span>
- <span data-ttu-id="2b419-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2b419-159">TabPrintPreview</span></span>
- <span data-ttu-id="2b419-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="2b419-160">TabMerge</span></span>
- <span data-ttu-id="2b419-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="2b419-161">TabGrayscale</span></span>
- <span data-ttu-id="2b419-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="2b419-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="2b419-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2b419-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="2b419-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="2b419-164">TabSlideMaster</span></span>
- <span data-ttu-id="2b419-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="2b419-165">TabHandoutMaster</span></span>
- <span data-ttu-id="2b419-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="2b419-166">TabNotesMaster</span></span>
- <span data-ttu-id="2b419-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2b419-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2b419-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="2b419-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="2b419-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="2b419-169">OneNote</span></span>

- <span data-ttu-id="2b419-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2b419-170">**TabHome**</span></span>
- <span data-ttu-id="2b419-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2b419-171">**TabInsert**</span></span>
- <span data-ttu-id="2b419-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2b419-172">**TabView**</span></span>
- <span data-ttu-id="2b419-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2b419-173">TabDeveloper</span></span>
- <span data-ttu-id="2b419-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2b419-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="2b419-175">Group</span><span class="sxs-lookup"><span data-stu-id="2b419-175">Group</span></span>

<span data-ttu-id="2b419-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="2b419-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="2b419-180">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="2b419-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
