---
title: Élément Group dans le fichier manifeste
description: Définit un groupe de contrôles d’interface utilisateur dans un onglet.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4717f6aeff3cd8ac34ee289252054417c489b89
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340462"
---
# <a name="group-element"></a>Élément Group

Définit un groupe de contrôles d’interface utilisateur dans un onglet. Sur les onglets personnalisés, le add-in peut créer plusieurs groupes. Les compléments sont limités à un onglet personnalisé.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Oui  | ID unique du groupe.|

### <a name="id-attribute"></a>Attribut id

Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique pour tous les éléments Group dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)      | Oui |  Étiquette d’un groupe.  |
|  [Icon](icon.md)      | Oui |  Image d’un groupe. Non pris en charge dans Outlook’autres. |
|  [Contrôle](#control)    | Non |  Représente un objet Control. Peut être zéro ou plus.  |
|  [OfficeControl](#officecontrol)  | Non | Représente l’un des contrôles Office intégrés. Peut être zéro ou plus. Non pris en charge dans Outlook’autres.|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le groupe doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. Non pris en charge dans Outlook’autres. |

### <a name="label"></a>Label

Obligatoire. Libellé du groupe. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).

### <a name="icon"></a>Icône

Obligatoire. Si un onglet contient un grand nombre de groupes et que la fenêtre du programme est resserée, l’image spécifiée peut s’afficher à la place.

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook les autres.

### <a name="control"></a>Contrôle

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un OfficeControl**. Pour plus d’informations sur les types de contrôles pris en charge, voir [l’élément](control.md) Control. L’ordre **des** contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon** .

```xml
<Group id="Contoso.CustomTab1.group1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button1">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **contrôle**. Incluez un ou plusieurs contrôles Office intégrés dans le groupe avec des `<OfficeControl>` éléments. L’attribut `id` spécifie l’ID du contrôle Office intégré. Pour trouver l’ID d’un contrôle, voir [Rechercher les ID des contrôles et des groupes de contrôles](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). L’ordre **des** contrôles **et OfficeControl** dans le manifeste est interchangeable et peut être entremêlé s’il existe plusieurs éléments, mais tous doivent se trouver sous l’élément **Icon** .

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook les autres.

```xml
<Group id="Contoso.CustomTab2.group2">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Facultatif (booléen). Spécifie si le groupe  sera masqué sur les combinaisons d’applications et de plateformes qui la prise en charge d’une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation. La valeur par défaut, si elle n’est pas présente, est `false`. S’il est utilisé, **OverriddenByRibbonApi** doit être le *premier* enfant de **Group**. Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook les autres.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.CustomTab3">
    <Group id="Contoso.CustomTab3.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
