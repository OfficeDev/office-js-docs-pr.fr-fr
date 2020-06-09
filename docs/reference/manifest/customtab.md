---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612303"
---
# <a name="customtab-element"></a><span data-ttu-id="ce296-103">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce296-103">CustomTab element</span></span>

<span data-ttu-id="ce296-104">Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="ce296-104">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="ce296-105">Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="ce296-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="ce296-p102">Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="ce296-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="ce296-109">L’attribut **ID** doit être unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="ce296-109">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ce296-110">Dans Outlook sur Mac, l' `CustomTab` élément n’est pas disponible et vous devez utiliser [OfficeTab](officetab.md) à la place.</span><span class="sxs-lookup"><span data-stu-id="ce296-110">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ce296-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="ce296-111">Child elements</span></span>

|  <span data-ttu-id="ce296-112">Élément</span><span class="sxs-lookup"><span data-stu-id="ce296-112">Element</span></span> |  <span data-ttu-id="ce296-113">Requis</span><span class="sxs-lookup"><span data-stu-id="ce296-113">Required</span></span>  |  <span data-ttu-id="ce296-114">Description</span><span class="sxs-lookup"><span data-stu-id="ce296-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ce296-115">Group</span><span class="sxs-lookup"><span data-stu-id="ce296-115">Group</span></span>](group.md)      | <span data-ttu-id="ce296-116">Oui</span><span class="sxs-lookup"><span data-stu-id="ce296-116">Yes</span></span> |  <span data-ttu-id="ce296-117">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="ce296-117">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="ce296-118">Label</span><span class="sxs-lookup"><span data-stu-id="ce296-118">Label</span></span>](#label-tab)      | <span data-ttu-id="ce296-119">Oui</span><span class="sxs-lookup"><span data-stu-id="ce296-119">Yes</span></span> |  <span data-ttu-id="ce296-120">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="ce296-120">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="ce296-121">Group</span><span class="sxs-lookup"><span data-stu-id="ce296-121">Group</span></span>

<span data-ttu-id="ce296-p103">Obligatoire. Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="ce296-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="ce296-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="ce296-124">Label (Tab)</span></span>

<span data-ttu-id="ce296-125">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="ce296-125">Required.</span></span> <span data-ttu-id="ce296-126">Étiquette de l’onglet personnalisé. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="ce296-126">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="ce296-127">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce296-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
