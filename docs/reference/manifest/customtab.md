---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087937"
---
# <a name="customtab-element"></a><span data-ttu-id="e69bd-103">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="e69bd-103">CustomTab element</span></span>

<span data-ttu-id="e69bd-104">Dans le ruban, spécifiez l’onglet et le groupe pour vos commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="e69bd-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="e69bd-105">Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="e69bd-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="e69bd-106">Dans les onglets personnalisés, le complément peut contenir des groupes personnalisés ou intégrés.</span><span class="sxs-lookup"><span data-stu-id="e69bd-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="e69bd-107">Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="e69bd-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="e69bd-108">L’attribut **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="e69bd-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e69bd-109">Dans Outlook sur Mac, l' `CustomTab` élément n’est pas disponible et vous devez utiliser [OfficeTab](officetab.md) à la place.</span><span class="sxs-lookup"><span data-stu-id="e69bd-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e69bd-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e69bd-110">Child elements</span></span>

|  <span data-ttu-id="e69bd-111">Élément</span><span class="sxs-lookup"><span data-stu-id="e69bd-111">Element</span></span> |  <span data-ttu-id="e69bd-112">Requis</span><span class="sxs-lookup"><span data-stu-id="e69bd-112">Required</span></span>  |  <span data-ttu-id="e69bd-113">Description</span><span class="sxs-lookup"><span data-stu-id="e69bd-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e69bd-114">Group</span><span class="sxs-lookup"><span data-stu-id="e69bd-114">Group</span></span>](group.md)      | <span data-ttu-id="e69bd-115">Non</span><span class="sxs-lookup"><span data-stu-id="e69bd-115">No</span></span> |  <span data-ttu-id="e69bd-116">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="e69bd-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="e69bd-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="e69bd-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="e69bd-118">Non</span><span class="sxs-lookup"><span data-stu-id="e69bd-118">No</span></span> |  <span data-ttu-id="e69bd-119">Représente un groupe de contrôles Office prédéfini.</span><span class="sxs-lookup"><span data-stu-id="e69bd-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="e69bd-120">Label</span><span class="sxs-lookup"><span data-stu-id="e69bd-120">Label</span></span>](#label-tab)      | <span data-ttu-id="e69bd-121">Oui</span><span class="sxs-lookup"><span data-stu-id="e69bd-121">Yes</span></span> |  <span data-ttu-id="e69bd-122">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="e69bd-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="e69bd-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="e69bd-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="e69bd-124">Non</span><span class="sxs-lookup"><span data-stu-id="e69bd-124">No</span></span> |  <span data-ttu-id="e69bd-125">Spécifie que l’onglet personnalisé doit se trouver immédiatement après un onglet Office prédéfini spécifié.</span><span class="sxs-lookup"><span data-stu-id="e69bd-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="e69bd-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="e69bd-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="e69bd-127">Non</span><span class="sxs-lookup"><span data-stu-id="e69bd-127">No</span></span> |  <span data-ttu-id="e69bd-128">Spécifie que l’onglet personnalisé doit se trouver immédiatement avant un onglet Office prédéfini spécifié.</span><span class="sxs-lookup"><span data-stu-id="e69bd-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="e69bd-129">Group</span><span class="sxs-lookup"><span data-stu-id="e69bd-129">Group</span></span>

<span data-ttu-id="e69bd-130">Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un élément **OfficeGroup** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="e69bd-131">Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="e69bd-131">See [Group element](group.md).</span></span> <span data-ttu-id="e69bd-132">L’ordre des **groupes** et des **OfficeGroup** dans le manifeste doit être l’ordre dans lequel vous souhaitez qu’ils apparaissent dans l’onglet personnalisé. Ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être au-dessus de l’élément **label** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="e69bd-133">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="e69bd-133">OfficeGroup</span></span>

<span data-ttu-id="e69bd-134">Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un élément **Group** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="e69bd-135">Représente un groupe de contrôles Office prédéfini.</span><span class="sxs-lookup"><span data-stu-id="e69bd-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="e69bd-136">L’attribut **ID** spécifie l’ID du groupe Office prédéfini.</span><span class="sxs-lookup"><span data-stu-id="e69bd-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="e69bd-137">Pour Rechercher l’ID d’un groupe prédéfini, voir [Rechercher les ID de contrôles et les groupes](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)de contrôles.</span><span class="sxs-lookup"><span data-stu-id="e69bd-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="e69bd-138">L’ordre des **groupes** et des **OfficeGroup** dans le manifeste doit être l’ordre dans lequel vous souhaitez qu’ils apparaissent dans l’onglet personnalisé. Ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être au-dessus de l’élément **label** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="e69bd-139">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="e69bd-139">Label (Tab)</span></span>

<span data-ttu-id="e69bd-140">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="e69bd-140">Required.</span></span> <span data-ttu-id="e69bd-141">Étiquette de l’onglet personnalisé. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="e69bd-141">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="e69bd-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="e69bd-142">InsertAfter</span></span>

<span data-ttu-id="e69bd-143">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="e69bd-143">Optional.</span></span> <span data-ttu-id="e69bd-144">Spécifie que l’onglet personnalisé doit se trouver immédiatement après un onglet Office prédéfini spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ».</span><span class="sxs-lookup"><span data-stu-id="e69bd-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="e69bd-145">(Voir [Rechercher les ID des contrôles et des groupes de](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)contrôles.) Le cas échéant, doit se trouver après l’élément **label** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="e69bd-146">Vous ne pouvez pas avoir à la fois **InsertAfter** et **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="e69bd-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="e69bd-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="e69bd-147">InsertBefore</span></span>

<span data-ttu-id="e69bd-148">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="e69bd-148">Optional.</span></span> <span data-ttu-id="e69bd-149">Spécifie que l’onglet personnalisé doit se trouver immédiatement avant un onglet Office prédéfini spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ».</span><span class="sxs-lookup"><span data-stu-id="e69bd-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="e69bd-150">(Voir [Rechercher les ID des contrôles et des groupes de](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)contrôles.)  Le cas échéant, doit se trouver après l’élément **label** .</span><span class="sxs-lookup"><span data-stu-id="e69bd-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="e69bd-151">Vous ne pouvez pas avoir à la fois **InsertAfter** et **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="e69bd-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="e69bd-152">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="e69bd-152">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
