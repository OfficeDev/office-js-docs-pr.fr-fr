---
title: Élément CustomTab dans le fichier manifest
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554012"
---
# <a name="customtab-element"></a>Élément CustomTab

Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (soit  **Accueil**,  **Message**, ou  **Réunion**), ou d’un onglet personnalisé défini par le complément.

Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.

L’attribut **id** doit être unique au sein du manifeste.

> [!IMPORTANT]
> Dans Outlook sur Mac, l' `CustomTab` élément n’est pas disponible et vous devez utiliser [OfficeTab](officetab.md) à la place.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Oui |  Définit un groupe de commandes.  |
|  [Label](#label-tab)      | Oui |  Étiquette pour CustomTab ou Group.  |

### <a name="group"></a>Group

Obligatoire. Voir [Élément group](group.md).

### <a name="label-tab"></a>Label (Tab)

Obligatoire. Étiquette de l’onglet personnalisé. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).


## <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
