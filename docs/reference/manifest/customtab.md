---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173926"
---
# <a name="customtab-element"></a><span data-ttu-id="0f13f-103">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="0f13f-103">CustomTab element</span></span>

<span data-ttu-id="0f13f-104">Dans le ruban, spécifiez l’onglet et le groupe pour vos commandes de module de recherche.</span><span class="sxs-lookup"><span data-stu-id="0f13f-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="0f13f-105">Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="0f13f-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="0f13f-106">Sur les onglets personnalisés, le add-in peut avoir des groupes personnalisés ou intégrés.</span><span class="sxs-lookup"><span data-stu-id="0f13f-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="0f13f-107">Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="0f13f-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="0f13f-108">**L’attribut id** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="0f13f-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f13f-109">Dans Outlook sur Mac, l’élément n’est pas disponible, vous devez `CustomTab` donc utiliser [OfficeTab](officetab.md) à la place.</span><span class="sxs-lookup"><span data-stu-id="0f13f-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0f13f-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="0f13f-110">Child elements</span></span>

|  <span data-ttu-id="0f13f-111">Élément</span><span class="sxs-lookup"><span data-stu-id="0f13f-111">Element</span></span> |  <span data-ttu-id="0f13f-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="0f13f-112">Required</span></span>  |  <span data-ttu-id="0f13f-113">Description</span><span class="sxs-lookup"><span data-stu-id="0f13f-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0f13f-114">Group</span><span class="sxs-lookup"><span data-stu-id="0f13f-114">Group</span></span>](group.md)      | <span data-ttu-id="0f13f-115">Non</span><span class="sxs-lookup"><span data-stu-id="0f13f-115">No</span></span> |  <span data-ttu-id="0f13f-116">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="0f13f-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="0f13f-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="0f13f-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="0f13f-118">Non</span><span class="sxs-lookup"><span data-stu-id="0f13f-118">No</span></span> |  <span data-ttu-id="0f13f-119">Représente un groupe de contrôles Office intégré.</span><span class="sxs-lookup"><span data-stu-id="0f13f-119">Represents a built-in Office control group.</span></span> <span data-ttu-id="0f13f-120">**Important**: non disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-120">**Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="0f13f-121">Label</span><span class="sxs-lookup"><span data-stu-id="0f13f-121">Label</span></span>](#label-tab)      | <span data-ttu-id="0f13f-122">Oui</span><span class="sxs-lookup"><span data-stu-id="0f13f-122">Yes</span></span> |  <span data-ttu-id="0f13f-123">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="0f13f-123">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="0f13f-124">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="0f13f-124">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="0f13f-125">Non</span><span class="sxs-lookup"><span data-stu-id="0f13f-125">No</span></span> |  <span data-ttu-id="0f13f-126">Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office intégré spécifié. **Important**: Non disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-126">Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="0f13f-127">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="0f13f-127">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="0f13f-128">Non</span><span class="sxs-lookup"><span data-stu-id="0f13f-128">No</span></span> |  <span data-ttu-id="0f13f-129">Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office intégré spécifié. **Important**: Non disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-129">Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="0f13f-130">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="0f13f-130">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="0f13f-131">Non</span><span class="sxs-lookup"><span data-stu-id="0f13f-131">No</span></span> |  <span data-ttu-id="0f13f-132">Spécifie si l’onglet personnalisé doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="0f13f-132">Specifies whether the custom tab should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="0f13f-133">**Important**: non disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-133">**Important**: Not available in Outlook.</span></span> |

### <a name="group"></a><span data-ttu-id="0f13f-134">Groupe</span><span class="sxs-lookup"><span data-stu-id="0f13f-134">Group</span></span>

<span data-ttu-id="0f13f-135">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un élément OfficeGroup.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-135">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="0f13f-136">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="0f13f-136">See [Group element](group.md).</span></span> <span data-ttu-id="0f13f-137">L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-137">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="0f13f-138">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="0f13f-138">OfficeGroup</span></span>

<span data-ttu-id="0f13f-139">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **élément Group.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-139">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="0f13f-140">Représente un groupe de contrôles Office intégré.</span><span class="sxs-lookup"><span data-stu-id="0f13f-140">Represents a built-in Office control group.</span></span> <span data-ttu-id="0f13f-141">**L’attribut id** spécifie l’ID du groupe Office intégré.</span><span class="sxs-lookup"><span data-stu-id="0f13f-141">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="0f13f-142">Pour trouver l’ID d’un groupe intégré, voir Rechercher les ID des contrôles et des [groupes de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="0f13f-142">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="0f13f-143">L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-143">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f13f-144">`OfficeGroup`L’élément n’est pas disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-144">The `OfficeGroup` element is not available in Outlook.</span></span>

### <a name="label-tab"></a><span data-ttu-id="0f13f-145">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="0f13f-145">Label (Tab)</span></span>

<span data-ttu-id="0f13f-146">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="0f13f-146">Required.</span></span> <span data-ttu-id="0f13f-147">Étiquette de l’onglet personnalisé. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="0f13f-147">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="0f13f-148">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="0f13f-148">InsertAfter</span></span>

<span data-ttu-id="0f13f-149">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="0f13f-149">Optional.</span></span> <span data-ttu-id="0f13f-150">Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office intégré spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ».</span><span class="sxs-lookup"><span data-stu-id="0f13f-150">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="0f13f-151">(Voir [Rechercher les ID des contrôles et des groupes de contrôles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) S’il est présent, il doit se trouver après **l’élément Label.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-151">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="0f13f-152">Vous ne pouvez pas **avoir à la fois InsertAfter** et **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="0f13f-152">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f13f-153">`InsertAfter`L’élément n’est pas disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-153">The `InsertAfter` element is not available in Outlook.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="0f13f-154">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="0f13f-154">InsertBefore</span></span>

<span data-ttu-id="0f13f-155">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="0f13f-155">Optional.</span></span> <span data-ttu-id="0f13f-156">Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office intégré spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ».</span><span class="sxs-lookup"><span data-stu-id="0f13f-156">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="0f13f-157">(Voir [Rechercher les ID des contrôles et des groupes de contrôles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  S’il est présent, il doit se trouver après **l’élément Label.**</span><span class="sxs-lookup"><span data-stu-id="0f13f-157">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="0f13f-158">Vous ne pouvez pas **avoir à la fois InsertAfter** et **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="0f13f-158">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f13f-159">`InsertBefore`L’élément n’est pas disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-159">The `InsertBefore` element is not available in Outlook.</span></span>

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="0f13f-160">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="0f13f-160">OverriddenByRibbonApi</span></span>

<span data-ttu-id="0f13f-161">Facultatif (booléen).</span><span class="sxs-lookup"><span data-stu-id="0f13f-161">Optional (boolean).</span></span> <span data-ttu-id="0f13f-162">Spécifie si **CustomTab** sera masqué sur les combinaisons d’applications et de plateformes qui la prise en charge d’une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation.</span><span class="sxs-lookup"><span data-stu-id="0f13f-162">Specifies whether the **CustomTab** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="0f13f-163">La valeur par défaut, si elle n’est pas présente, est `false` .</span><span class="sxs-lookup"><span data-stu-id="0f13f-163">The default value, if not present, is `false`.</span></span> <span data-ttu-id="0f13f-164">S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="0f13f-164">If used, **OverriddenByRibbonApi** must be the *first* child of **CustomTab**.</span></span> <span data-ttu-id="0f13f-165">Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="0f13f-165">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f13f-166">`OverriddenByRibbonApi`L’élément n’est pas disponible dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f13f-166">The `OverriddenByRibbonApi` element is not available in Outlook.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="0f13f-167">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="0f13f-167">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
