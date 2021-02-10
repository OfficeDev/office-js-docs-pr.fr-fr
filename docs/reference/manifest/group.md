---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173961"
---
# <a name="group-element"></a><span data-ttu-id="54fd2-103">Élément Group</span><span class="sxs-lookup"><span data-stu-id="54fd2-103">Group element</span></span>

<span data-ttu-id="54fd2-104">Définit un groupe de contrôles d’interface utilisateur dans un onglet. Sur les onglets personnalisés, le add-in peut créer plusieurs groupes.</span><span class="sxs-lookup"><span data-stu-id="54fd2-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="54fd2-105">Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="54fd2-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="54fd2-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="54fd2-106">Attributes</span></span>

|  <span data-ttu-id="54fd2-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="54fd2-107">Attribute</span></span>  |  <span data-ttu-id="54fd2-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="54fd2-108">Required</span></span>  |  <span data-ttu-id="54fd2-109">Description</span><span class="sxs-lookup"><span data-stu-id="54fd2-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="54fd2-110">id</span><span class="sxs-lookup"><span data-stu-id="54fd2-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="54fd2-111">Oui</span><span class="sxs-lookup"><span data-stu-id="54fd2-111">Yes</span></span>  | <span data-ttu-id="54fd2-112">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="54fd2-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="54fd2-113">Attribut id</span><span class="sxs-lookup"><span data-stu-id="54fd2-113">id attribute</span></span>

<span data-ttu-id="54fd2-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="54fd2-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="54fd2-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="54fd2-118">Child elements</span></span>

|  <span data-ttu-id="54fd2-119">Élément</span><span class="sxs-lookup"><span data-stu-id="54fd2-119">Element</span></span> |  <span data-ttu-id="54fd2-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="54fd2-120">Required</span></span>  |  <span data-ttu-id="54fd2-121">Description</span><span class="sxs-lookup"><span data-stu-id="54fd2-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="54fd2-122">Label</span><span class="sxs-lookup"><span data-stu-id="54fd2-122">Label</span></span>](#label)      | <span data-ttu-id="54fd2-123">Oui</span><span class="sxs-lookup"><span data-stu-id="54fd2-123">Yes</span></span> |  <span data-ttu-id="54fd2-124">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="54fd2-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="54fd2-125">Icon</span><span class="sxs-lookup"><span data-stu-id="54fd2-125">Icon</span></span>](icon.md)      | <span data-ttu-id="54fd2-126">Oui</span><span class="sxs-lookup"><span data-stu-id="54fd2-126">Yes</span></span> |  <span data-ttu-id="54fd2-127">Image d’un groupe.</span><span class="sxs-lookup"><span data-stu-id="54fd2-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="54fd2-128">Contrôle</span><span class="sxs-lookup"><span data-stu-id="54fd2-128">Control</span></span>](#control)    | <span data-ttu-id="54fd2-129">Non</span><span class="sxs-lookup"><span data-stu-id="54fd2-129">No</span></span> |  <span data-ttu-id="54fd2-130">Représente un objet Control.</span><span class="sxs-lookup"><span data-stu-id="54fd2-130">Represents a Control object.</span></span> <span data-ttu-id="54fd2-131">Peut être zéro ou plus.</span><span class="sxs-lookup"><span data-stu-id="54fd2-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="54fd2-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="54fd2-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="54fd2-133">Non</span><span class="sxs-lookup"><span data-stu-id="54fd2-133">No</span></span> | <span data-ttu-id="54fd2-134">Représente l’un des contrôles Office intégrés.</span><span class="sxs-lookup"><span data-stu-id="54fd2-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="54fd2-135">Peut être zéro ou plus.</span><span class="sxs-lookup"><span data-stu-id="54fd2-135">Can be zero or more.</span></span> |
|  [<span data-ttu-id="54fd2-136">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="54fd2-136">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="54fd2-137">Non</span><span class="sxs-lookup"><span data-stu-id="54fd2-137">No</span></span> |  <span data-ttu-id="54fd2-138">Spécifie si le groupe doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="54fd2-138">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span>  |

### <a name="label"></a><span data-ttu-id="54fd2-139">Label</span><span class="sxs-lookup"><span data-stu-id="54fd2-139">Label</span></span>

<span data-ttu-id="54fd2-140">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="54fd2-140">Required.</span></span> <span data-ttu-id="54fd2-141">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="54fd2-141">The label of the group.</span></span> <span data-ttu-id="54fd2-142">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="54fd2-142">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="54fd2-143">Icône</span><span class="sxs-lookup"><span data-stu-id="54fd2-143">Icon</span></span>

<span data-ttu-id="54fd2-144">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="54fd2-144">Required.</span></span> <span data-ttu-id="54fd2-145">Si un onglet contient un grand nombre de groupes et que la fenêtre de programme est re resserée, l’image spécifiée peut s’afficher à la place.</span><span class="sxs-lookup"><span data-stu-id="54fd2-145">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="54fd2-146">Contrôle</span><span class="sxs-lookup"><span data-stu-id="54fd2-146">Control</span></span>

<span data-ttu-id="54fd2-147">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="54fd2-147">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="54fd2-148">Pour plus d’informations sur les types de contrôles pris en charge, voir [l’élément](control.md) Control.</span><span class="sxs-lookup"><span data-stu-id="54fd2-148">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="54fd2-149">L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**</span><span class="sxs-lookup"><span data-stu-id="54fd2-149">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="54fd2-150">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="54fd2-150">OfficeControl</span></span>

<span data-ttu-id="54fd2-151">Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **contrôle**.</span><span class="sxs-lookup"><span data-stu-id="54fd2-151">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="54fd2-152">Inclure un ou plusieurs contrôles Office intégrés dans le groupe avec des `<OfficeControl>` éléments.</span><span class="sxs-lookup"><span data-stu-id="54fd2-152">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="54fd2-153">`id`L’attribut spécifie l’ID du contrôle Office intégré.</span><span class="sxs-lookup"><span data-stu-id="54fd2-153">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="54fd2-154">Pour trouver l’ID d’un contrôle, voir Rechercher les ID des contrôles et des groupes [de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="54fd2-154">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="54fd2-155">L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**</span><span class="sxs-lookup"><span data-stu-id="54fd2-155">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="54fd2-156">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="54fd2-156">OverriddenByRibbonApi</span></span>

<span data-ttu-id="54fd2-157">Facultatif (booléen).</span><span class="sxs-lookup"><span data-stu-id="54fd2-157">Optional (boolean).</span></span> <span data-ttu-id="54fd2-158">Spécifie si  le groupe sera masqué sur les combinaisons d’applications et de plateformes qui prisent en charge une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation.</span><span class="sxs-lookup"><span data-stu-id="54fd2-158">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="54fd2-159">La valeur par défaut, si elle n’est pas présente, est `false` .</span><span class="sxs-lookup"><span data-stu-id="54fd2-159">The default value, if not present, is `false`.</span></span> <span data-ttu-id="54fd2-160">S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **Group**.</span><span class="sxs-lookup"><span data-stu-id="54fd2-160">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="54fd2-161">Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="54fd2-161">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

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
