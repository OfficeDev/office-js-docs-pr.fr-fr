# <a name="officetab-element"></a><span data-ttu-id="47b25-101">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="47b25-101">OfficeTab element</span></span>

<span data-ttu-id="47b25-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="47b25-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="47b25-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="47b25-105">Child elements</span></span>

|  <span data-ttu-id="47b25-106">Élément</span><span class="sxs-lookup"><span data-stu-id="47b25-106">Element</span></span> |  <span data-ttu-id="47b25-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="47b25-107">Required</span></span>  |  <span data-ttu-id="47b25-108">Description</span><span class="sxs-lookup"><span data-stu-id="47b25-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="47b25-109">Groupe</span><span class="sxs-lookup"><span data-stu-id="47b25-109">Group</span></span>      | <span data-ttu-id="47b25-110">Oui</span><span class="sxs-lookup"><span data-stu-id="47b25-110">Yes</span></span> |  <span data-ttu-id="47b25-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="47b25-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="47b25-113">Les valeurs suivantes sont des valeurs `id` d’onglet valides par hôte.</span><span class="sxs-lookup"><span data-stu-id="47b25-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="47b25-114">Les valeurs en **gras** sont prises en charge à la fois sur le bureau et en ligne (par exemple, Word 2016 ou versions ultérieures pour Windows et Word Online).</span><span class="sxs-lookup"><span data-stu-id="47b25-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="47b25-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="47b25-115">Outlook</span></span>

- <span data-ttu-id="47b25-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="47b25-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="47b25-117">Word</span><span class="sxs-lookup"><span data-stu-id="47b25-117">Word</span></span>

- <span data-ttu-id="47b25-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="47b25-118">**TabHome**</span></span>
- <span data-ttu-id="47b25-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="47b25-119">**TabInsert**</span></span>
- <span data-ttu-id="47b25-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="47b25-120">TabWordDesign</span></span>
- <span data-ttu-id="47b25-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="47b25-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="47b25-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="47b25-122">TabReferences</span></span>
- <span data-ttu-id="47b25-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="47b25-123">TabMailings</span></span>
- <span data-ttu-id="47b25-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="47b25-124">TabReviewWord</span></span>
- <span data-ttu-id="47b25-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="47b25-125">**TabView**</span></span>
- <span data-ttu-id="47b25-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="47b25-126">TabDeveloper</span></span>
- <span data-ttu-id="47b25-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="47b25-127">TabAddIns</span></span>
- <span data-ttu-id="47b25-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="47b25-128">TabBlogPost</span></span>
- <span data-ttu-id="47b25-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="47b25-129">TabBlogInsert</span></span>
- <span data-ttu-id="47b25-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="47b25-130">TabPrintPreview</span></span>
- <span data-ttu-id="47b25-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="47b25-131">TabOutlining</span></span>
- <span data-ttu-id="47b25-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="47b25-132">TabConflicts</span></span>
- <span data-ttu-id="47b25-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="47b25-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="47b25-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="47b25-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="47b25-135">Excel</span><span class="sxs-lookup"><span data-stu-id="47b25-135">Excel</span></span>

- <span data-ttu-id="47b25-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="47b25-136">**TabHome**</span></span>
- <span data-ttu-id="47b25-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="47b25-137">**TabInsert**</span></span>
- <span data-ttu-id="47b25-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="47b25-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="47b25-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="47b25-139">TabFormulas</span></span>
- <span data-ttu-id="47b25-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="47b25-140">**TabData**</span></span>
- <span data-ttu-id="47b25-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="47b25-141">**TabReview**</span></span>
- <span data-ttu-id="47b25-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="47b25-142">**TabView**</span></span>
- <span data-ttu-id="47b25-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="47b25-143">TabDeveloper</span></span>
- <span data-ttu-id="47b25-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="47b25-144">TabAddIns</span></span>
- <span data-ttu-id="47b25-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="47b25-145">TabPrintPreview</span></span>
- <span data-ttu-id="47b25-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="47b25-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="47b25-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="47b25-147">PowerPoint</span></span>

- <span data-ttu-id="47b25-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="47b25-148">**TabHome**</span></span>
- <span data-ttu-id="47b25-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="47b25-149">**TabInsert**</span></span>
- <span data-ttu-id="47b25-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="47b25-150">**TabDesign**</span></span>
- <span data-ttu-id="47b25-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="47b25-151">**TabTransitions**</span></span>
- <span data-ttu-id="47b25-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="47b25-152">**TabAnimations**</span></span>
- <span data-ttu-id="47b25-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="47b25-153">TabSlideShow</span></span>
- <span data-ttu-id="47b25-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="47b25-154">TabReview</span></span>
- <span data-ttu-id="47b25-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="47b25-155">**TabView**</span></span>
- <span data-ttu-id="47b25-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="47b25-156">TabDeveloper</span></span>
- <span data-ttu-id="47b25-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="47b25-157">TabAddIns</span></span>
- <span data-ttu-id="47b25-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="47b25-158">TabPrintPreview</span></span>
- <span data-ttu-id="47b25-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="47b25-159">TabMerge</span></span>
- <span data-ttu-id="47b25-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="47b25-160">TabGrayscale</span></span>
- <span data-ttu-id="47b25-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="47b25-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="47b25-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="47b25-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="47b25-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="47b25-163">TabSlideMaster</span></span>
- <span data-ttu-id="47b25-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="47b25-164">TabHandoutMaster</span></span>
- <span data-ttu-id="47b25-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="47b25-165">TabNotesMaster</span></span>
- <span data-ttu-id="47b25-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="47b25-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="47b25-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="47b25-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="47b25-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="47b25-168">OneNote</span></span>

- <span data-ttu-id="47b25-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="47b25-169">**TabHome**</span></span>
- <span data-ttu-id="47b25-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="47b25-170">**TabInsert**</span></span>
- <span data-ttu-id="47b25-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="47b25-171">**TabView**</span></span>
- <span data-ttu-id="47b25-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="47b25-172">TabDeveloper</span></span>
- <span data-ttu-id="47b25-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="47b25-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="47b25-174">Groupe</span><span class="sxs-lookup"><span data-stu-id="47b25-174">Group</span></span>

<span data-ttu-id="47b25-p104">|||UNTRANSLATED_CONTENT_START|||A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="47b25-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="47b25-179">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="47b25-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
