---
title: Élément CustomTab dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1c3c6883a1feb94299feb35c078431e6e2e322c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450631"
---
# <a name="customtab-element"></a><span data-ttu-id="e3f3a-102">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="e3f3a-102">CustomTab element</span></span>

<span data-ttu-id="e3f3a-p101">Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (soit  **Accueil**,  **Message**, ou  **Réunion**), ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="e3f3a-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="e3f3a-p102">Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="e3f3a-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="e3f3a-108">L’attribut **id** doit être unique au sein du manifeste.</span><span class="sxs-lookup"><span data-stu-id="e3f3a-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e3f3a-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e3f3a-109">Child elements</span></span>

|  <span data-ttu-id="e3f3a-110">Élément</span><span class="sxs-lookup"><span data-stu-id="e3f3a-110">Element</span></span> |  <span data-ttu-id="e3f3a-111">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e3f3a-111">Required</span></span>  |  <span data-ttu-id="e3f3a-112">Description</span><span class="sxs-lookup"><span data-stu-id="e3f3a-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e3f3a-113">Group</span><span class="sxs-lookup"><span data-stu-id="e3f3a-113">Group</span></span>](group.md)      | <span data-ttu-id="e3f3a-114">Oui</span><span class="sxs-lookup"><span data-stu-id="e3f3a-114">Yes</span></span> |  <span data-ttu-id="e3f3a-115">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="e3f3a-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="e3f3a-116">Label</span><span class="sxs-lookup"><span data-stu-id="e3f3a-116">Label</span></span>](#label-tab)      | <span data-ttu-id="e3f3a-117">Oui</span><span class="sxs-lookup"><span data-stu-id="e3f3a-117">Yes</span></span> |  <span data-ttu-id="e3f3a-118">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="e3f3a-118">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="e3f3a-119">Control</span><span class="sxs-lookup"><span data-stu-id="e3f3a-119">Control</span></span>](control.md)    | <span data-ttu-id="e3f3a-120">Oui</span><span class="sxs-lookup"><span data-stu-id="e3f3a-120">Yes</span></span> |  <span data-ttu-id="e3f3a-121">Ensemble d’un ou de plusieurs objets Control</span><span class="sxs-lookup"><span data-stu-id="e3f3a-121">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="e3f3a-122">Group</span><span class="sxs-lookup"><span data-stu-id="e3f3a-122">Group</span></span>

<span data-ttu-id="e3f3a-p103">Obligatoire. Voir [Élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="e3f3a-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="e3f3a-125">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="e3f3a-125">Label (Tab)</span></span>

<span data-ttu-id="e3f3a-p104">Obligatoire. Étiquette de l’onglet personnalisé. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e3f3a-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="e3f3a-128">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="e3f3a-128">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
