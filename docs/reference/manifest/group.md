---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 6fe07497e98bd77aad7ad296850a0b9f9e9bf9a4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718180"
---
# <a name="group-element"></a><span data-ttu-id="d6a03-103">Élément Group</span><span class="sxs-lookup"><span data-stu-id="d6a03-103">Group element</span></span>

<span data-ttu-id="d6a03-p101">Définit un groupe de contrôles d’interface utilisateur dans un onglet.  Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="d6a03-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="d6a03-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="d6a03-107">Attributes</span></span>

|  <span data-ttu-id="d6a03-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="d6a03-108">Attribute</span></span>  |  <span data-ttu-id="d6a03-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d6a03-109">Required</span></span>  |  <span data-ttu-id="d6a03-110">Description</span><span class="sxs-lookup"><span data-stu-id="d6a03-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d6a03-111">id</span><span class="sxs-lookup"><span data-stu-id="d6a03-111">id</span></span>](#id-attribute)  |  <span data-ttu-id="d6a03-112">Oui</span><span class="sxs-lookup"><span data-stu-id="d6a03-112">Yes</span></span>  | <span data-ttu-id="d6a03-113">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="d6a03-113">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="d6a03-114">Attribut id</span><span class="sxs-lookup"><span data-stu-id="d6a03-114">id attribute</span></span>

<span data-ttu-id="d6a03-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="d6a03-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d6a03-119">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d6a03-119">Child elements</span></span>
|  <span data-ttu-id="d6a03-120">Élément</span><span class="sxs-lookup"><span data-stu-id="d6a03-120">Element</span></span> |  <span data-ttu-id="d6a03-121">Requis</span><span class="sxs-lookup"><span data-stu-id="d6a03-121">Required</span></span>  |  <span data-ttu-id="d6a03-122">Description</span><span class="sxs-lookup"><span data-stu-id="d6a03-122">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d6a03-123">Label</span><span class="sxs-lookup"><span data-stu-id="d6a03-123">Label</span></span>](#label)      | <span data-ttu-id="d6a03-124">Oui</span><span class="sxs-lookup"><span data-stu-id="d6a03-124">Yes</span></span> |  <span data-ttu-id="d6a03-125">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="d6a03-125">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="d6a03-126">Icon</span><span class="sxs-lookup"><span data-stu-id="d6a03-126">Icon</span></span>](icon.md)      | <span data-ttu-id="d6a03-127">Oui</span><span class="sxs-lookup"><span data-stu-id="d6a03-127">Yes</span></span> |  <span data-ttu-id="d6a03-128">Image d’un groupe.</span><span class="sxs-lookup"><span data-stu-id="d6a03-128">The image for a group.</span></span>  |
|  [<span data-ttu-id="d6a03-129">Control</span><span class="sxs-lookup"><span data-stu-id="d6a03-129">Control</span></span>](#control)    | <span data-ttu-id="d6a03-130">Oui</span><span class="sxs-lookup"><span data-stu-id="d6a03-130">Yes</span></span> |  <span data-ttu-id="d6a03-131">Ensemble d’un ou de plusieurs objets Control.</span><span class="sxs-lookup"><span data-stu-id="d6a03-131">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="d6a03-132">Label</span><span class="sxs-lookup"><span data-stu-id="d6a03-132">Label</span></span> 

<span data-ttu-id="d6a03-133">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="d6a03-133">Required.</span></span> <span data-ttu-id="d6a03-134">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="d6a03-134">The label of the group.</span></span> <span data-ttu-id="d6a03-135">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="d6a03-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="d6a03-136">Icône</span><span class="sxs-lookup"><span data-stu-id="d6a03-136">Icon</span></span>

<span data-ttu-id="d6a03-137">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="d6a03-137">Required.</span></span> <span data-ttu-id="d6a03-138">Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est redimensionnée, l’image spécifiée peut s’afficher à la place.</span><span class="sxs-lookup"><span data-stu-id="d6a03-138">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="d6a03-139">Contrôle</span><span class="sxs-lookup"><span data-stu-id="d6a03-139">Control</span></span>
<span data-ttu-id="d6a03-140">Un groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="d6a03-140">A group requires at least one control.</span></span> <span data-ttu-id="d6a03-141">Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="d6a03-141">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
