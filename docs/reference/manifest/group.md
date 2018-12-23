---
title: Élément Group dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 13cd9bbe6f602fd1779caea487e34177c3e9d483
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433696"
---
# <a name="group-element"></a><span data-ttu-id="bcda1-102">Élément Group</span><span class="sxs-lookup"><span data-stu-id="bcda1-102">Group element</span></span>

<span data-ttu-id="bcda1-p101">Définit un groupe de contrôles d’interface utilisateur dans un onglet.  Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="bcda1-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="bcda1-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="bcda1-106">Attributes</span></span>

|  <span data-ttu-id="bcda1-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="bcda1-107">Attribute</span></span>  |  <span data-ttu-id="bcda1-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="bcda1-108">Required</span></span>  |  <span data-ttu-id="bcda1-109">Description</span><span class="sxs-lookup"><span data-stu-id="bcda1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bcda1-110">id</span><span class="sxs-lookup"><span data-stu-id="bcda1-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="bcda1-111">Oui</span><span class="sxs-lookup"><span data-stu-id="bcda1-111">Yes</span></span>  | <span data-ttu-id="bcda1-112">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="bcda1-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="bcda1-113">Attribut id</span><span class="sxs-lookup"><span data-stu-id="bcda1-113">id attribute</span></span>

<span data-ttu-id="bcda1-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="bcda1-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bcda1-118">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="bcda1-118">Child elements</span></span>
|  <span data-ttu-id="bcda1-119">Élément</span><span class="sxs-lookup"><span data-stu-id="bcda1-119">Element</span></span> |  <span data-ttu-id="bcda1-120">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="bcda1-120">Required</span></span>  |  <span data-ttu-id="bcda1-121">Description</span><span class="sxs-lookup"><span data-stu-id="bcda1-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bcda1-122">Label</span><span class="sxs-lookup"><span data-stu-id="bcda1-122">Label</span></span>](#label)      | <span data-ttu-id="bcda1-123">Oui</span><span class="sxs-lookup"><span data-stu-id="bcda1-123">Yes</span></span> |  <span data-ttu-id="bcda1-124">Étiquette pour CustomTab ou group.</span><span class="sxs-lookup"><span data-stu-id="bcda1-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="bcda1-125">Control</span><span class="sxs-lookup"><span data-stu-id="bcda1-125">Control</span></span>](#control)    | <span data-ttu-id="bcda1-126">Oui</span><span class="sxs-lookup"><span data-stu-id="bcda1-126">Yes</span></span> |  <span data-ttu-id="bcda1-127">Ensemble d’un ou de plusieurs objets Control.</span><span class="sxs-lookup"><span data-stu-id="bcda1-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="bcda1-128">Étiquette</span><span class="sxs-lookup"><span data-stu-id="bcda1-128">Label</span></span> 

<span data-ttu-id="bcda1-p103">Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bcda1-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="bcda1-132">Contrôle</span><span class="sxs-lookup"><span data-stu-id="bcda1-132">Control</span></span>
<span data-ttu-id="bcda1-133">Un groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="bcda1-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```