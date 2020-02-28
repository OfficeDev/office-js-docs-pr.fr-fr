---
title: Élément CustomTab dans le fichier manifest
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: ba0419b6cf9cc4a0c1e3038dbb7f972e65868ec4
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323804"
---
# <a name="customtab-element"></a>Élément CustomTab

Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.

Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.

L’attribut **ID** doit être unique dans le manifeste.

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

Obligatoire. Étiquette de l’onglet personnalisé. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .


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
