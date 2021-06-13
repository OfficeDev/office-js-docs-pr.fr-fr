---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 89ed16f7996ab06bd21e1ebaa71c959b11af2029
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893511"
---
# <a name="group-element"></a><span data-ttu-id="7a3fc-103">Élément Group</span><span class="sxs-lookup"><span data-stu-id="7a3fc-103">Group element</span></span>

<span data-ttu-id="7a3fc-104">Définit un groupe de contrôles d’interface utilisateur dans un onglet. Sur les onglets personnalisés, le add-in peut créer plusieurs groupes.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="7a3fc-105">Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="7a3fc-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="7a3fc-106">Attributes</span></span>

|  <span data-ttu-id="7a3fc-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="7a3fc-107">Attribute</span></span>  |  <span data-ttu-id="7a3fc-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7a3fc-108">Required</span></span>  |  <span data-ttu-id="7a3fc-109">Description</span><span class="sxs-lookup"><span data-stu-id="7a3fc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7a3fc-110">id</span><span class="sxs-lookup"><span data-stu-id="7a3fc-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="7a3fc-111">Oui</span><span class="sxs-lookup"><span data-stu-id="7a3fc-111">Yes</span></span>  | <span data-ttu-id="7a3fc-112">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="7a3fc-113">Attribut id</span><span class="sxs-lookup"><span data-stu-id="7a3fc-113">id attribute</span></span>

<span data-ttu-id="7a3fc-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7a3fc-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7a3fc-118">Child elements</span></span>

|  <span data-ttu-id="7a3fc-119">Élément</span><span class="sxs-lookup"><span data-stu-id="7a3fc-119">Element</span></span> |  <span data-ttu-id="7a3fc-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7a3fc-120">Required</span></span>  |  <span data-ttu-id="7a3fc-121">Description</span><span class="sxs-lookup"><span data-stu-id="7a3fc-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7a3fc-122">Label</span><span class="sxs-lookup"><span data-stu-id="7a3fc-122">Label</span></span>](#label)      | <span data-ttu-id="7a3fc-123">Oui</span><span class="sxs-lookup"><span data-stu-id="7a3fc-123">Yes</span></span> |  <span data-ttu-id="7a3fc-124">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="7a3fc-125">Icon</span><span class="sxs-lookup"><span data-stu-id="7a3fc-125">Icon</span></span>](icon.md)      | <span data-ttu-id="7a3fc-126">Oui</span><span class="sxs-lookup"><span data-stu-id="7a3fc-126">Yes</span></span> |  <span data-ttu-id="7a3fc-127">Image d’un groupe.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-127">The image for a group.</span></span> <span data-ttu-id="7a3fc-128">Non pris en charge dans Outlook des modules.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-128">Not supported in Outlook add-ins.</span></span> |
|  [<span data-ttu-id="7a3fc-129">Contrôle</span><span class="sxs-lookup"><span data-stu-id="7a3fc-129">Control</span></span>](#control)    | <span data-ttu-id="7a3fc-130">Non</span><span class="sxs-lookup"><span data-stu-id="7a3fc-130">No</span></span> |  <span data-ttu-id="7a3fc-131">Représente un objet Control.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-131">Represents a Control object.</span></span> <span data-ttu-id="7a3fc-132">Peut être zéro ou plus.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-132">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="7a3fc-133">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="7a3fc-133">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="7a3fc-134">Non</span><span class="sxs-lookup"><span data-stu-id="7a3fc-134">No</span></span> | <span data-ttu-id="7a3fc-135">Représente l’un des contrôles Office intégrés.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-135">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="7a3fc-136">Peut être zéro ou plus.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-136">Can be zero or more.</span></span> <span data-ttu-id="7a3fc-137">Non pris en charge dans Outlook des modules.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-137">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="7a3fc-138">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="7a3fc-138">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="7a3fc-139">Non</span><span class="sxs-lookup"><span data-stu-id="7a3fc-139">No</span></span> |  <span data-ttu-id="7a3fc-140">Spécifie si le groupe doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-140">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="7a3fc-141">Non pris en charge dans Outlook des modules.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-141">Not supported in Outlook add-ins.</span></span> |

### <a name="label"></a><span data-ttu-id="7a3fc-142">Label</span><span class="sxs-lookup"><span data-stu-id="7a3fc-142">Label</span></span>

<span data-ttu-id="7a3fc-143">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-143">Required.</span></span> <span data-ttu-id="7a3fc-144">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-144">The label of the group.</span></span> <span data-ttu-id="7a3fc-145">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="7a3fc-145">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="7a3fc-146">Icône</span><span class="sxs-lookup"><span data-stu-id="7a3fc-146">Icon</span></span>

<span data-ttu-id="7a3fc-147">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-147">Required.</span></span> <span data-ttu-id="7a3fc-148">Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est re resserée, l’image spécifiée peut s’afficher à la place.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-148">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

> [!NOTE]
> <span data-ttu-id="7a3fc-149">Cet élément enfant n’est pas pris en charge dans Outlook de développement.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-149">This child element is not supported in Outlook add-ins.</span></span>

### <a name="control"></a><span data-ttu-id="7a3fc-150">Contrôle</span><span class="sxs-lookup"><span data-stu-id="7a3fc-150">Control</span></span>

<span data-ttu-id="7a3fc-151">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-151">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="7a3fc-152">Pour plus d’informations sur les types de contrôles pris en charge, voir [l’élément](control.md) Control.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-152">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="7a3fc-153">L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**</span><span class="sxs-lookup"><span data-stu-id="7a3fc-153">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a><span data-ttu-id="7a3fc-154">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="7a3fc-154">OfficeControl</span></span>

<span data-ttu-id="7a3fc-155">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **contrôle**.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-155">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="7a3fc-156">Incluez un ou plusieurs contrôles Office intégrés dans le groupe avec des `<OfficeControl>` éléments.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-156">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="7a3fc-157">L’attribut spécifie l’ID du contrôle Office `id` intégré.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-157">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="7a3fc-158">Pour trouver l’ID d’un contrôle, voir Rechercher les ID des contrôles et des groupes [de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="7a3fc-158">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="7a3fc-159">L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**</span><span class="sxs-lookup"><span data-stu-id="7a3fc-159">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

> [!NOTE]
> <span data-ttu-id="7a3fc-160">Cet élément enfant n’est pas pris en charge dans Outlook de développement.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-160">This child element is not supported in Outlook add-ins.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="7a3fc-161">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="7a3fc-161">OverriddenByRibbonApi</span></span>

<span data-ttu-id="7a3fc-162">Facultatif (booléen).</span><span class="sxs-lookup"><span data-stu-id="7a3fc-162">Optional (boolean).</span></span> <span data-ttu-id="7a3fc-163">Spécifie si  le groupe sera masqué sur les combinaisons d’applications et de plateformes qui supportent une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-163">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="7a3fc-164">La valeur par défaut, si elle n’est pas présente, est `false` .</span><span class="sxs-lookup"><span data-stu-id="7a3fc-164">The default value, if not present, is `false`.</span></span> <span data-ttu-id="7a3fc-165">S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **Group**.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-165">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="7a3fc-166">Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="7a3fc-166">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!NOTE]
> <span data-ttu-id="7a3fc-167">Cet élément enfant n’est pas pris en charge dans Outlook de développement.</span><span class="sxs-lookup"><span data-stu-id="7a3fc-167">This child element is not supported in Outlook add-ins.</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
