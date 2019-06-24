---
title: Élément OfficeTab dans le fichier manifest
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d073d712cec2fd58e957ffe8f344d7443d1e896e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127561"
---
# <a name="officetab-element"></a><span data-ttu-id="f4290-102">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="f4290-102">OfficeTab element</span></span>

<span data-ttu-id="f4290-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="f4290-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f4290-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="f4290-106">Child elements</span></span>

|  <span data-ttu-id="f4290-107">Élément</span><span class="sxs-lookup"><span data-stu-id="f4290-107">Element</span></span> |  <span data-ttu-id="f4290-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="f4290-108">Required</span></span>  |  <span data-ttu-id="f4290-109">Description</span><span class="sxs-lookup"><span data-stu-id="f4290-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f4290-110">Groupe</span><span class="sxs-lookup"><span data-stu-id="f4290-110">Group</span></span>      | <span data-ttu-id="f4290-111">Oui</span><span class="sxs-lookup"><span data-stu-id="f4290-111">Yes</span></span> |  <span data-ttu-id="f4290-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="f4290-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="f4290-114">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="f4290-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="f4290-115">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).</span><span class="sxs-lookup"><span data-stu-id="f4290-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="f4290-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="f4290-116">Outlook</span></span>

- <span data-ttu-id="f4290-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="f4290-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="f4290-118">Word</span><span class="sxs-lookup"><span data-stu-id="f4290-118">Word</span></span>

- <span data-ttu-id="f4290-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4290-119">**TabHome**</span></span>
- <span data-ttu-id="f4290-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4290-120">**TabInsert**</span></span>
- <span data-ttu-id="f4290-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="f4290-121">TabWordDesign</span></span>
- <span data-ttu-id="f4290-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="f4290-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="f4290-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="f4290-123">TabReferences</span></span>
- <span data-ttu-id="f4290-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="f4290-124">TabMailings</span></span>
- <span data-ttu-id="f4290-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="f4290-125">TabReviewWord</span></span>
- <span data-ttu-id="f4290-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4290-126">**TabView**</span></span>
- <span data-ttu-id="f4290-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4290-127">TabDeveloper</span></span>
- <span data-ttu-id="f4290-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4290-128">TabAddIns</span></span>
- <span data-ttu-id="f4290-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="f4290-129">TabBlogPost</span></span>
- <span data-ttu-id="f4290-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="f4290-130">TabBlogInsert</span></span>
- <span data-ttu-id="f4290-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4290-131">TabPrintPreview</span></span>
- <span data-ttu-id="f4290-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="f4290-132">TabOutlining</span></span>
- <span data-ttu-id="f4290-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="f4290-133">TabConflicts</span></span>
- <span data-ttu-id="f4290-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4290-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="f4290-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="f4290-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="f4290-136">Excel</span><span class="sxs-lookup"><span data-stu-id="f4290-136">Excel</span></span>

- <span data-ttu-id="f4290-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4290-137">**TabHome**</span></span>
- <span data-ttu-id="f4290-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4290-138">**TabInsert**</span></span>
- <span data-ttu-id="f4290-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="f4290-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="f4290-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="f4290-140">TabFormulas</span></span>
- <span data-ttu-id="f4290-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="f4290-141">**TabData**</span></span>
- <span data-ttu-id="f4290-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="f4290-142">**TabReview**</span></span>
- <span data-ttu-id="f4290-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4290-143">**TabView**</span></span>
- <span data-ttu-id="f4290-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4290-144">TabDeveloper</span></span>
- <span data-ttu-id="f4290-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4290-145">TabAddIns</span></span>
- <span data-ttu-id="f4290-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4290-146">TabPrintPreview</span></span>
- <span data-ttu-id="f4290-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4290-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="f4290-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f4290-148">PowerPoint</span></span>

- <span data-ttu-id="f4290-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4290-149">**TabHome**</span></span>
- <span data-ttu-id="f4290-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4290-150">**TabInsert**</span></span>
- <span data-ttu-id="f4290-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="f4290-151">**TabDesign**</span></span>
- <span data-ttu-id="f4290-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="f4290-152">**TabTransitions**</span></span>
- <span data-ttu-id="f4290-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="f4290-153">**TabAnimations**</span></span>
- <span data-ttu-id="f4290-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="f4290-154">TabSlideShow</span></span>
- <span data-ttu-id="f4290-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="f4290-155">TabReview</span></span>
- <span data-ttu-id="f4290-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4290-156">**TabView**</span></span>
- <span data-ttu-id="f4290-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4290-157">TabDeveloper</span></span>
- <span data-ttu-id="f4290-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4290-158">TabAddIns</span></span>
- <span data-ttu-id="f4290-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4290-159">TabPrintPreview</span></span>
- <span data-ttu-id="f4290-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="f4290-160">TabMerge</span></span>
- <span data-ttu-id="f4290-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="f4290-161">TabGrayscale</span></span>
- <span data-ttu-id="f4290-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="f4290-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="f4290-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="f4290-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="f4290-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="f4290-164">TabSlideMaster</span></span>
- <span data-ttu-id="f4290-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="f4290-165">TabHandoutMaster</span></span>
- <span data-ttu-id="f4290-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="f4290-166">TabNotesMaster</span></span>
- <span data-ttu-id="f4290-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4290-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="f4290-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="f4290-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="f4290-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="f4290-169">OneNote</span></span>

- <span data-ttu-id="f4290-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4290-170">**TabHome**</span></span>
- <span data-ttu-id="f4290-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4290-171">**TabInsert**</span></span>
- <span data-ttu-id="f4290-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4290-172">**TabView**</span></span>
- <span data-ttu-id="f4290-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4290-173">TabDeveloper</span></span>
- <span data-ttu-id="f4290-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4290-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="f4290-175">Group</span><span class="sxs-lookup"><span data-stu-id="f4290-175">Group</span></span>

<span data-ttu-id="f4290-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="f4290-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="f4290-180">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="f4290-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
