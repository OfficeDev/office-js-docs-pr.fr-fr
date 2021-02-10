---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173961"
---
# <a name="group-element"></a>Élément Group

Définit un groupe de contrôles d’interface utilisateur dans un onglet. Sur les onglets personnalisés, le add-in peut créer plusieurs groupes. Les compléments sont limités à un onglet personnalisé.

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
|  [Contrôle](#control)    | Non |  Représente un objet Control. Peut être zéro ou plus.  |
|  [OfficeControl](#officecontrol)  | Non | Représente l’un des contrôles Office intégrés. Peut être zéro ou plus. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le groupe doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés.  |

### <a name="label"></a>Label

Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)

### <a name="icon"></a>Icône

Obligatoire. Si un onglet contient un grand nombre de groupes et que la fenêtre de programme est re resserée, l’image spécifiée peut s’afficher à la place.

### <a name="control"></a>Contrôle

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un OfficeControl**. Pour plus d’informations sur les types de contrôles pris en charge, voir [l’élément](control.md) Control. L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **contrôle**. Inclure un ou plusieurs contrôles Office intégrés dans le groupe avec des `<OfficeControl>` éléments. `id`L’attribut spécifie l’ID du contrôle Office intégré. Pour trouver l’ID d’un contrôle, voir Rechercher les ID des contrôles et des groupes [de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Facultatif (booléen). Spécifie si  le groupe sera masqué sur les combinaisons d’applications et de plateformes qui prisent en charge une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation. La valeur par défaut, si elle n’est pas présente, est `false` . S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **Group**. Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
