---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087944"
---
# <a name="group-element"></a>Élément Group

Définit un groupe de contrôles d’interface utilisateur dans un onglet. Dans les onglets personnalisés, le complément peut créer plusieurs groupes. Les compléments sont limités à un onglet personnalisé.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Oui  | ID unique du groupe.|

### <a name="id-attribute"></a>Attribut id

Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)      | Oui |  Étiquette pour CustomTab ou group.  |
|  [Icon](icon.md)      | Oui |  Image d’un groupe.  |
|  [Contrôle](#control)    | Non |  Représente un objet Control. Peut être zéro ou plusieurs.  |
|  [OfficeControl](#officecontrol)  | Non | Représente l’un des contrôles Office prédéfinis. Peut être zéro ou plusieurs. |

### <a name="label"></a>Label

Obligatoire. Libellé du groupe. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .

### <a name="icon"></a>Icône

Obligatoire. Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est redimensionnée, l’image spécifiée peut s’afficher à la place.

### <a name="control"></a>Contrôle

Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un **OfficeControl**. Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) . L’ordre des **contrôles** et **OfficeControl** dans le manifeste est interchangeable et ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être sous l’élément **Icon** .

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

### <a name="officecontrol"></a>OfficeControl

Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un **contrôle**. Inclure un ou plusieurs contrôles Office prédéfinis dans le groupe avec des `<OfficeControl>` éléments. L' `id` attribut spécifie l’ID du contrôle Office prédéfini. Pour Rechercher l’ID d’un contrôle, voir [Rechercher les ID des contrôles et des groupes](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)de contrôles. L’ordre des **contrôles** et **OfficeControl** dans le manifeste est interchangeable et ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être sous l’élément **Icon** .

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
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```
