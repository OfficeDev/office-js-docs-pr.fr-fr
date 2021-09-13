---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 06/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 09260ab52910235ab63149769cc989ffbda03ffb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153572"
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
|  [Icon](icon.md)      | Oui |  Image d’un groupe. Non pris en charge dans Outlook des modules. |
|  [Contrôle](#control)    | Non |  Représente un objet Control. Peut être zéro ou plus.  |
|  [OfficeControl](#officecontrol)  | Non | Représente l’un des contrôles Office intégrés. Peut être zéro ou plus. Non pris en charge dans Outlook des modules.|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le groupe doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. Non pris en charge dans Outlook des modules. |

### <a name="label"></a>Label

Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)

### <a name="icon"></a>Icône

Obligatoire. Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est re resserée, l’image spécifiée peut s’afficher à la place.

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook de développement.

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

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **contrôle**. Incluez un ou plusieurs contrôles Office intégrés dans le groupe avec des `<OfficeControl>` éléments. L’attribut spécifie l’ID du contrôle Office `id` intégré. Pour trouver l’ID d’un contrôle, voir Rechercher les ID des contrôles et des groupes [de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) L’ordre  des contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon.**

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook de développement.

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

Facultatif (booléen). Spécifie si  le groupe sera masqué sur les combinaisons d’applications et de plateformes qui supportent une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation. La valeur par défaut, si elle n’est pas présente, est `false` . S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **Group**. Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook de développement.

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
