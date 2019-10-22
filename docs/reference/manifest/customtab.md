---
title: Élément CustomTab dans le fichier manifest
description: ''
ms.date: 04/29/2019
localization_priority: Normal
ms.openlocfilehash: 4fa7dd86736b5ab421be5653f2e256a6b84fb480
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33517393"
---
# <a name="customtab-element"></a><span data-ttu-id="6786f-102">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="6786f-102">CustomTab element</span></span>

<span data-ttu-id="6786f-p101">Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (soit  **Accueil**,  **Message**, ou  **Réunion**), ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="6786f-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="6786f-p102">Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="6786f-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="6786f-108">L’attribut **id** doit être unique au sein du manifeste.</span><span class="sxs-lookup"><span data-stu-id="6786f-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6786f-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="6786f-109">Child elements</span></span>

|  <span data-ttu-id="6786f-110">Élément</span><span class="sxs-lookup"><span data-stu-id="6786f-110">Element</span></span> |  <span data-ttu-id="6786f-111">Requis</span><span class="sxs-lookup"><span data-stu-id="6786f-111">Required</span></span>  |  <span data-ttu-id="6786f-112">Description</span><span class="sxs-lookup"><span data-stu-id="6786f-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6786f-113">Group</span><span class="sxs-lookup"><span data-stu-id="6786f-113">Group</span></span>](group.md)      | <span data-ttu-id="6786f-114">Oui</span><span class="sxs-lookup"><span data-stu-id="6786f-114">Yes</span></span> |  <span data-ttu-id="6786f-115">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="6786f-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="6786f-116">Label</span><span class="sxs-lookup"><span data-stu-id="6786f-116">Label</span></span>](#label-tab)      | <span data-ttu-id="6786f-117">Oui</span><span class="sxs-lookup"><span data-stu-id="6786f-117">Yes</span></span> |  <span data-ttu-id="6786f-118">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="6786f-118">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="6786f-119">Group</span><span class="sxs-lookup"><span data-stu-id="6786f-119">Group</span></span>

<span data-ttu-id="6786f-p103">Obligatoire. Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="6786f-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="6786f-122">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="6786f-122">Label (Tab)</span></span>

<span data-ttu-id="6786f-p104">Obligatoire. Étiquette de l’onglet personnalisé. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="6786f-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="6786f-125">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="6786f-125">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
