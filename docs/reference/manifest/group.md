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
# <a name="group-element"></a>Élément Group

Définit un groupe de contrôles d’interface utilisateur dans un onglet.  Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Oui  | ID unique du groupe.|

### <a name="id-attribute"></a>Attribut id

Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.

## <a name="child-elements"></a>Éléments enfants
|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)      | Oui |  Étiquette pour CustomTab ou group.  |
|  [Icon](icon.md)      | Oui |  Image d’un groupe.  |
|  [Control](#control)    | Oui |  Ensemble d’un ou de plusieurs objets Control.  |

### <a name="label"></a>Label 

Obligatoire. Libellé du groupe. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .

### <a name="icon"></a>Icône

Obligatoire. Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est redimensionnée, l’image spécifiée peut s’afficher à la place.

### <a name="control"></a>Contrôle
Un groupe requiert au moins un contrôle. Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) .

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
