---
title: Élément OfficeTab dans le fichier manifest
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 1bf9f1d1e08a8147b52f93923229ef8fb8556fcf
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952270"
---
# <a name="officetab-element"></a><span data-ttu-id="a6cbd-102">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a6cbd-102">OfficeTab element</span></span>

<span data-ttu-id="a6cbd-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="a6cbd-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a6cbd-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="a6cbd-106">Child elements</span></span>

|  <span data-ttu-id="a6cbd-107">Élément</span><span class="sxs-lookup"><span data-stu-id="a6cbd-107">Element</span></span> |  <span data-ttu-id="a6cbd-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a6cbd-108">Required</span></span>  |  <span data-ttu-id="a6cbd-109">Description</span><span class="sxs-lookup"><span data-stu-id="a6cbd-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a6cbd-110">Groupe</span><span class="sxs-lookup"><span data-stu-id="a6cbd-110">Group</span></span>      | <span data-ttu-id="a6cbd-111">Oui</span><span class="sxs-lookup"><span data-stu-id="a6cbd-111">Yes</span></span> |  <span data-ttu-id="a6cbd-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="a6cbd-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="a6cbd-114">Les valeurs suivantes sont des valeurs `id` d’onglet valides par l’hôte.</span><span class="sxs-lookup"><span data-stu-id="a6cbd-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="a6cbd-115">Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word Online).</span><span class="sxs-lookup"><span data-stu-id="a6cbd-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="a6cbd-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="a6cbd-116">Outlook</span></span>

- <span data-ttu-id="a6cbd-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="a6cbd-118">Word</span><span class="sxs-lookup"><span data-stu-id="a6cbd-118">Word</span></span>

- <span data-ttu-id="a6cbd-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-119">**TabHome**</span></span>
- <span data-ttu-id="a6cbd-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-120">**TabInsert**</span></span>
- <span data-ttu-id="a6cbd-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="a6cbd-121">TabWordDesign</span></span>
- <span data-ttu-id="a6cbd-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="a6cbd-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="a6cbd-123">TabReferences</span></span>
- <span data-ttu-id="a6cbd-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="a6cbd-124">TabMailings</span></span>
- <span data-ttu-id="a6cbd-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="a6cbd-125">TabReviewWord</span></span>
- <span data-ttu-id="a6cbd-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-126">**TabView**</span></span>
- <span data-ttu-id="a6cbd-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="a6cbd-127">TabDeveloper</span></span>
- <span data-ttu-id="a6cbd-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="a6cbd-128">TabAddIns</span></span>
- <span data-ttu-id="a6cbd-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="a6cbd-129">TabBlogPost</span></span>
- <span data-ttu-id="a6cbd-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="a6cbd-130">TabBlogInsert</span></span>
- <span data-ttu-id="a6cbd-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="a6cbd-131">TabPrintPreview</span></span>
- <span data-ttu-id="a6cbd-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="a6cbd-132">TabOutlining</span></span>
- <span data-ttu-id="a6cbd-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="a6cbd-133">TabConflicts</span></span>
- <span data-ttu-id="a6cbd-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="a6cbd-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="a6cbd-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="a6cbd-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="a6cbd-136">Excel</span><span class="sxs-lookup"><span data-stu-id="a6cbd-136">Excel</span></span>

- <span data-ttu-id="a6cbd-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-137">**TabHome**</span></span>
- <span data-ttu-id="a6cbd-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-138">**TabInsert**</span></span>
- <span data-ttu-id="a6cbd-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="a6cbd-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="a6cbd-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="a6cbd-140">TabFormulas</span></span>
- <span data-ttu-id="a6cbd-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-141">**TabData**</span></span>
- <span data-ttu-id="a6cbd-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-142">**TabReview**</span></span>
- <span data-ttu-id="a6cbd-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-143">**TabView**</span></span>
- <span data-ttu-id="a6cbd-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="a6cbd-144">TabDeveloper</span></span>
- <span data-ttu-id="a6cbd-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="a6cbd-145">TabAddIns</span></span>
- <span data-ttu-id="a6cbd-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="a6cbd-146">TabPrintPreview</span></span>
- <span data-ttu-id="a6cbd-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="a6cbd-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="a6cbd-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a6cbd-148">PowerPoint</span></span>

- <span data-ttu-id="a6cbd-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-149">**TabHome**</span></span>
- <span data-ttu-id="a6cbd-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-150">**TabInsert**</span></span>
- <span data-ttu-id="a6cbd-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-151">**TabDesign**</span></span>
- <span data-ttu-id="a6cbd-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-152">**TabTransitions**</span></span>
- <span data-ttu-id="a6cbd-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-153">**TabAnimations**</span></span>
- <span data-ttu-id="a6cbd-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="a6cbd-154">TabSlideShow</span></span>
- <span data-ttu-id="a6cbd-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="a6cbd-155">TabReview</span></span>
- <span data-ttu-id="a6cbd-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-156">**TabView**</span></span>
- <span data-ttu-id="a6cbd-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="a6cbd-157">TabDeveloper</span></span>
- <span data-ttu-id="a6cbd-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="a6cbd-158">TabAddIns</span></span>
- <span data-ttu-id="a6cbd-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="a6cbd-159">TabPrintPreview</span></span>
- <span data-ttu-id="a6cbd-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="a6cbd-160">TabMerge</span></span>
- <span data-ttu-id="a6cbd-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="a6cbd-161">TabGrayscale</span></span>
- <span data-ttu-id="a6cbd-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="a6cbd-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="a6cbd-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="a6cbd-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="a6cbd-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="a6cbd-164">TabSlideMaster</span></span>
- <span data-ttu-id="a6cbd-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="a6cbd-165">TabHandoutMaster</span></span>
- <span data-ttu-id="a6cbd-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="a6cbd-166">TabNotesMaster</span></span>
- <span data-ttu-id="a6cbd-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="a6cbd-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="a6cbd-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="a6cbd-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="a6cbd-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="a6cbd-169">OneNote</span></span>

- <span data-ttu-id="a6cbd-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-170">**TabHome**</span></span>
- <span data-ttu-id="a6cbd-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-171">**TabInsert**</span></span>
- <span data-ttu-id="a6cbd-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="a6cbd-172">**TabView**</span></span>
- <span data-ttu-id="a6cbd-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="a6cbd-173">TabDeveloper</span></span>
- <span data-ttu-id="a6cbd-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="a6cbd-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="a6cbd-175">Group</span><span class="sxs-lookup"><span data-stu-id="a6cbd-175">Group</span></span>

<span data-ttu-id="a6cbd-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="a6cbd-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="a6cbd-180">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a6cbd-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
