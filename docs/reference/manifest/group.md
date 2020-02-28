---
title: Élément Group dans le fichier manifeste
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 27a168ea17352482e955e7a0d1f8267c7d6b17d8
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324861"
---
# <a name="group-element"></a><span data-ttu-id="40323-102">Élément Group</span><span class="sxs-lookup"><span data-stu-id="40323-102">Group element</span></span>

<span data-ttu-id="40323-p101">Définit un groupe de contrôles d’interface utilisateur dans un onglet.  Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="40323-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="40323-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="40323-106">Attributes</span></span>

|  <span data-ttu-id="40323-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="40323-107">Attribute</span></span>  |  <span data-ttu-id="40323-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="40323-108">Required</span></span>  |  <span data-ttu-id="40323-109">Description</span><span class="sxs-lookup"><span data-stu-id="40323-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="40323-110">id</span><span class="sxs-lookup"><span data-stu-id="40323-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="40323-111">Oui</span><span class="sxs-lookup"><span data-stu-id="40323-111">Yes</span></span>  | <span data-ttu-id="40323-112">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="40323-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="40323-113">Attribut id</span><span class="sxs-lookup"><span data-stu-id="40323-113">id attribute</span></span>

<span data-ttu-id="40323-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="40323-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="40323-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="40323-118">Child elements</span></span>
|  <span data-ttu-id="40323-119">Élément</span><span class="sxs-lookup"><span data-stu-id="40323-119">Element</span></span> |  <span data-ttu-id="40323-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="40323-120">Required</span></span>  |  <span data-ttu-id="40323-121">Description</span><span class="sxs-lookup"><span data-stu-id="40323-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="40323-122">Label</span><span class="sxs-lookup"><span data-stu-id="40323-122">Label</span></span>](#label)      | <span data-ttu-id="40323-123">Oui</span><span class="sxs-lookup"><span data-stu-id="40323-123">Yes</span></span> |  <span data-ttu-id="40323-124">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="40323-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="40323-125">Icon</span><span class="sxs-lookup"><span data-stu-id="40323-125">Icon</span></span>](icon.md)      | <span data-ttu-id="40323-126">Oui</span><span class="sxs-lookup"><span data-stu-id="40323-126">Yes</span></span> |  <span data-ttu-id="40323-127">Image d’un groupe.</span><span class="sxs-lookup"><span data-stu-id="40323-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="40323-128">Control</span><span class="sxs-lookup"><span data-stu-id="40323-128">Control</span></span>](#control)    | <span data-ttu-id="40323-129">Oui</span><span class="sxs-lookup"><span data-stu-id="40323-129">Yes</span></span> |  <span data-ttu-id="40323-130">Ensemble d’un ou de plusieurs objets Control.</span><span class="sxs-lookup"><span data-stu-id="40323-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="40323-131">Label</span><span class="sxs-lookup"><span data-stu-id="40323-131">Label</span></span> 

<span data-ttu-id="40323-132">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="40323-132">Required.</span></span> <span data-ttu-id="40323-133">Libellé du groupe.</span><span class="sxs-lookup"><span data-stu-id="40323-133">The label of the group.</span></span> <span data-ttu-id="40323-134">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="40323-134">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="40323-135">Icône</span><span class="sxs-lookup"><span data-stu-id="40323-135">Icon</span></span>

<span data-ttu-id="40323-136">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="40323-136">Required.</span></span> <span data-ttu-id="40323-137">Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est redimensionnée, l’image spécifiée peut s’afficher à la place.</span><span class="sxs-lookup"><span data-stu-id="40323-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="40323-138">Contrôle</span><span class="sxs-lookup"><span data-stu-id="40323-138">Control</span></span>
<span data-ttu-id="40323-139">Un groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="40323-139">A group requires at least one control.</span></span> <span data-ttu-id="40323-140">Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="40323-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
