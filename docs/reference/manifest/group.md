---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087944"
---
# <a name="group-element"></a><span data-ttu-id="0d444-103">Élément Group</span><span class="sxs-lookup"><span data-stu-id="0d444-103">Group element</span></span>

<span data-ttu-id="0d444-104">Définit un groupe de contrôles d’interface utilisateur dans un onglet. Dans les onglets personnalisés, le complément peut créer plusieurs groupes.</span><span class="sxs-lookup"><span data-stu-id="0d444-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="0d444-105">Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="0d444-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="0d444-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="0d444-106">Attributes</span></span>

|  <span data-ttu-id="0d444-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="0d444-107">Attribute</span></span>  |  <span data-ttu-id="0d444-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="0d444-108">Required</span></span>  |  <span data-ttu-id="0d444-109">Description</span><span class="sxs-lookup"><span data-stu-id="0d444-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0d444-110">id</span><span class="sxs-lookup"><span data-stu-id="0d444-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="0d444-111">Oui</span><span class="sxs-lookup"><span data-stu-id="0d444-111">Yes</span></span>  | <span data-ttu-id="0d444-112">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="0d444-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="0d444-113">Attribut id</span><span class="sxs-lookup"><span data-stu-id="0d444-113">id attribute</span></span>

<span data-ttu-id="0d444-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="0d444-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0d444-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="0d444-118">Child elements</span></span>

|  <span data-ttu-id="0d444-119">Élément</span><span class="sxs-lookup"><span data-stu-id="0d444-119">Element</span></span> |  <span data-ttu-id="0d444-120">Requis</span><span class="sxs-lookup"><span data-stu-id="0d444-120">Required</span></span>  |  <span data-ttu-id="0d444-121">Description</span><span class="sxs-lookup"><span data-stu-id="0d444-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0d444-122">Label</span><span class="sxs-lookup"><span data-stu-id="0d444-122">Label</span></span>](#label)      | <span data-ttu-id="0d444-123">Oui</span><span class="sxs-lookup"><span data-stu-id="0d444-123">Yes</span></span> |  <span data-ttu-id="0d444-124">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="0d444-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="0d444-125">Icon</span><span class="sxs-lookup"><span data-stu-id="0d444-125">Icon</span></span>](icon.md)      | <span data-ttu-id="0d444-126">Oui</span><span class="sxs-lookup"><span data-stu-id="0d444-126">Yes</span></span> |  <span data-ttu-id="0d444-127">Image d’un groupe.</span><span class="sxs-lookup"><span data-stu-id="0d444-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="0d444-128">Contrôle</span><span class="sxs-lookup"><span data-stu-id="0d444-128">Control</span></span>](#control)    | <span data-ttu-id="0d444-129">Non</span><span class="sxs-lookup"><span data-stu-id="0d444-129">No</span></span> |  <span data-ttu-id="0d444-130">Représente un objet Control.</span><span class="sxs-lookup"><span data-stu-id="0d444-130">Represents a Control object.</span></span> <span data-ttu-id="0d444-131">Peut être zéro ou plusieurs.</span><span class="sxs-lookup"><span data-stu-id="0d444-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="0d444-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="0d444-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="0d444-133">Non</span><span class="sxs-lookup"><span data-stu-id="0d444-133">No</span></span> | <span data-ttu-id="0d444-134">Représente l’un des contrôles Office prédéfinis.</span><span class="sxs-lookup"><span data-stu-id="0d444-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="0d444-135">Peut être zéro ou plusieurs.</span><span class="sxs-lookup"><span data-stu-id="0d444-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="0d444-136">Label</span><span class="sxs-lookup"><span data-stu-id="0d444-136">Label</span></span>

<span data-ttu-id="0d444-137">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="0d444-137">Required.</span></span> <span data-ttu-id="0d444-138">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="0d444-138">The label of the group.</span></span> <span data-ttu-id="0d444-139">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="0d444-139">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="0d444-140">Icône</span><span class="sxs-lookup"><span data-stu-id="0d444-140">Icon</span></span>

<span data-ttu-id="0d444-141">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="0d444-141">Required.</span></span> <span data-ttu-id="0d444-142">Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est redimensionnée, l’image spécifiée peut s’afficher à la place.</span><span class="sxs-lookup"><span data-stu-id="0d444-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="0d444-143">Contrôle</span><span class="sxs-lookup"><span data-stu-id="0d444-143">Control</span></span>

<span data-ttu-id="0d444-144">Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un **OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="0d444-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="0d444-145">Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="0d444-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="0d444-146">L’ordre des **contrôles** et **OfficeControl** dans le manifeste est interchangeable et ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être sous l’élément **Icon** .</span><span class="sxs-lookup"><span data-stu-id="0d444-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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

### <a name="officecontrol"></a><span data-ttu-id="0d444-147">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="0d444-147">OfficeControl</span></span>

<span data-ttu-id="0d444-148">Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un **contrôle**.</span><span class="sxs-lookup"><span data-stu-id="0d444-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="0d444-149">Inclure un ou plusieurs contrôles Office prédéfinis dans le groupe avec des `<OfficeControl>` éléments.</span><span class="sxs-lookup"><span data-stu-id="0d444-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="0d444-150">L' `id` attribut spécifie l’ID du contrôle Office prédéfini.</span><span class="sxs-lookup"><span data-stu-id="0d444-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="0d444-151">Pour Rechercher l’ID d’un contrôle, voir [Rechercher les ID des contrôles et des groupes](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)de contrôles.</span><span class="sxs-lookup"><span data-stu-id="0d444-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="0d444-152">L’ordre des **contrôles** et **OfficeControl** dans le manifeste est interchangeable et ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être sous l’élément **Icon** .</span><span class="sxs-lookup"><span data-stu-id="0d444-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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
