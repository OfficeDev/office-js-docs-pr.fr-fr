---
title: Élément OfficeTab dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b61c245c000f8bf13eb71c991ec57a125993c2fc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450463"
---
# <a name="officetab-element"></a><span data-ttu-id="671ea-102">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="671ea-102">OfficeTab element</span></span>

<span data-ttu-id="671ea-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="671ea-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="671ea-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="671ea-106">Child elements</span></span>

|  <span data-ttu-id="671ea-107">Élément</span><span class="sxs-lookup"><span data-stu-id="671ea-107">Element</span></span> |  <span data-ttu-id="671ea-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="671ea-108">Required</span></span>  |  <span data-ttu-id="671ea-109">Description</span><span class="sxs-lookup"><span data-stu-id="671ea-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="671ea-110">Groupe</span><span class="sxs-lookup"><span data-stu-id="671ea-110">Group</span></span>      | <span data-ttu-id="671ea-111">Oui</span><span class="sxs-lookup"><span data-stu-id="671ea-111">Yes</span></span> |  <span data-ttu-id="671ea-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="671ea-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="671ea-114">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="671ea-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="671ea-115">Les valeurs en **gras** sont prises en charge à la fois sur le bureau et en ligne (par exemple, Word 2016 pour Windows et Word Online).</span><span class="sxs-lookup"><span data-stu-id="671ea-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="671ea-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="671ea-116">Outlook</span></span>

- <span data-ttu-id="671ea-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="671ea-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="671ea-118">Word</span><span class="sxs-lookup"><span data-stu-id="671ea-118">Word</span></span>

- <span data-ttu-id="671ea-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="671ea-119">**TabHome**</span></span>
- <span data-ttu-id="671ea-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="671ea-120">**TabInsert**</span></span>
- <span data-ttu-id="671ea-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="671ea-121">TabWordDesign</span></span>
- <span data-ttu-id="671ea-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="671ea-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="671ea-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="671ea-123">TabReferences</span></span>
- <span data-ttu-id="671ea-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="671ea-124">TabMailings</span></span>
- <span data-ttu-id="671ea-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="671ea-125">TabReviewWord</span></span>
- <span data-ttu-id="671ea-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="671ea-126">**TabView**</span></span>
- <span data-ttu-id="671ea-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="671ea-127">TabDeveloper</span></span>
- <span data-ttu-id="671ea-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="671ea-128">TabAddIns</span></span>
- <span data-ttu-id="671ea-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="671ea-129">TabBlogPost</span></span>
- <span data-ttu-id="671ea-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="671ea-130">TabBlogInsert</span></span>
- <span data-ttu-id="671ea-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="671ea-131">TabPrintPreview</span></span>
- <span data-ttu-id="671ea-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="671ea-132">TabOutlining</span></span>
- <span data-ttu-id="671ea-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="671ea-133">TabConflicts</span></span>
- <span data-ttu-id="671ea-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="671ea-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="671ea-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="671ea-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="671ea-136">Excel</span><span class="sxs-lookup"><span data-stu-id="671ea-136">Excel</span></span>

- <span data-ttu-id="671ea-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="671ea-137">**TabHome**</span></span>
- <span data-ttu-id="671ea-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="671ea-138">**TabInsert**</span></span>
- <span data-ttu-id="671ea-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="671ea-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="671ea-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="671ea-140">TabFormulas</span></span>
- <span data-ttu-id="671ea-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="671ea-141">**TabData**</span></span>
- <span data-ttu-id="671ea-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="671ea-142">**TabReview**</span></span>
- <span data-ttu-id="671ea-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="671ea-143">**TabView**</span></span>
- <span data-ttu-id="671ea-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="671ea-144">TabDeveloper</span></span>
- <span data-ttu-id="671ea-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="671ea-145">TabAddIns</span></span>
- <span data-ttu-id="671ea-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="671ea-146">TabPrintPreview</span></span>
- <span data-ttu-id="671ea-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="671ea-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="671ea-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="671ea-148">PowerPoint</span></span>

- <span data-ttu-id="671ea-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="671ea-149">**TabHome**</span></span>
- <span data-ttu-id="671ea-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="671ea-150">**TabInsert**</span></span>
- <span data-ttu-id="671ea-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="671ea-151">**TabDesign**</span></span>
- <span data-ttu-id="671ea-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="671ea-152">**TabTransitions**</span></span>
- <span data-ttu-id="671ea-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="671ea-153">**TabAnimations**</span></span>
- <span data-ttu-id="671ea-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="671ea-154">TabSlideShow</span></span>
- <span data-ttu-id="671ea-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="671ea-155">TabReview</span></span>
- <span data-ttu-id="671ea-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="671ea-156">**TabView**</span></span>
- <span data-ttu-id="671ea-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="671ea-157">TabDeveloper</span></span>
- <span data-ttu-id="671ea-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="671ea-158">TabAddIns</span></span>
- <span data-ttu-id="671ea-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="671ea-159">TabPrintPreview</span></span>
- <span data-ttu-id="671ea-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="671ea-160">TabMerge</span></span>
- <span data-ttu-id="671ea-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="671ea-161">TabGrayscale</span></span>
- <span data-ttu-id="671ea-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="671ea-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="671ea-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="671ea-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="671ea-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="671ea-164">TabSlideMaster</span></span>
- <span data-ttu-id="671ea-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="671ea-165">TabHandoutMaster</span></span>
- <span data-ttu-id="671ea-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="671ea-166">TabNotesMaster</span></span>
- <span data-ttu-id="671ea-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="671ea-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="671ea-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="671ea-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="671ea-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="671ea-169">OneNote</span></span>

- <span data-ttu-id="671ea-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="671ea-170">**TabHome**</span></span>
- <span data-ttu-id="671ea-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="671ea-171">**TabInsert**</span></span>
- <span data-ttu-id="671ea-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="671ea-172">**TabView**</span></span>
- <span data-ttu-id="671ea-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="671ea-173">TabDeveloper</span></span>
- <span data-ttu-id="671ea-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="671ea-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="671ea-175">Group</span><span class="sxs-lookup"><span data-stu-id="671ea-175">Group</span></span>

<span data-ttu-id="671ea-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="671ea-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="671ea-180">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="671ea-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
