---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6a9540fd7e98464681a90021a36f7a7529186f7f
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340112"
---
# <a name="customtab-element"></a>Élément CustomTab

Définit un onglet personnalisé pour le ruban Office personnalisé. Ajoutez des contrôles et des groupes de ruban pour le add-in à l’un des onglets Office de build ou à votre propre onglet personnalisé. Utilisez l’élément **CustomTab** pour ajouter un onglet personnalisé au ruban. Sur les onglets personnalisés, le add-in peut avoir des groupes personnalisés ou intégrés. Les compléments sont limités à un onglet personnalisé.

> [!IMPORTANT]
> Dans Outlook Mac, l’élément **CustomTab** n’est pas disponible, mais vous pouvez placer des  groupes personnalisés de contrôles sur l’un des [Contrôles OfficeTab](officetab.md) intégrés à la place. Vous ne pouvez *pas placer des groupes intégrés* sur *des onglets* intégrés Outlook sur n’importe quelle plateforme.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

> [!NOTE]
> Certains éléments enfants ne sont pas valides dans les schémas de messagerie. Voir [éléments enfants](#child-elements).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). Requis par certains éléments enfants. Voir [éléments enfants](#child-elements).

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Oui  | ID unique de l’onglet personnalisé.|

### <a name="id-attribute"></a>Attribut id

Obligatoire. Identificateur unique de l’onglet personnalisé. Il s’agit d’une chaîne de 125 caractères au maximum. Il doit être unique dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Non |  Définit un groupe de commandes.  |
|  [OfficeGroup](#officegroup)      | Non |  Représente un groupe de contrôle Office intégré. **Important** : non disponible dans Outlook. |
|  [Label](#label-tab)      | Oui |  Étiquette de CustomTab.  |
|  [InsertAfter](#insertafter)      | Non |  Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office spécifié. **Important** : disponible uniquement dans PowerPoint. |
|  [InsertBefore](#insertbefore)      | Non |  Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office spécifié. **Important** : disponible uniquement dans PowerPoint. |

### <a name="group"></a>Group

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un élément OfficeGroup** . Voir [Élément group](group.md). L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label** .

### <a name="officegroup"></a>OfficeGroup

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **élément Group** . Représente un groupe de contrôle Office intégré. **L’attribut id** spécifie l’ID du groupe Office intégré. Pour trouver l’ID d’un groupe intégré, voir [Rechercher les ID des contrôles et des groupes de contrôles](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label** .

> [!IMPORTANT]
> **L’élément OfficeGroup** n’est pas disponible Outlook. Dans PowerPoint, il est en prévisualisation pour Mac et Windows, mais il est disponible pour les macros de production dans PowerPoint sur le web.

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="label-tab"></a>Label (Tab)

Obligatoire. Étiquette de l’onglet personnalisé. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md) .

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertafter"></a>InsertAfter

Facultatif. Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que `TabHome` ou `TabReview`.  Pour obtenir la liste des onglets intégrés, voir [OfficeTab](officetab.md). S’il est présent, il doit se trouver après **l’élément Label** . Vous ne pouvez pas avoir **à la fois InsertAfter** **et InsertBefore**.

> [!IMPORTANT]
> **L’élément InsertAfter** est disponible uniquement dans PowerPoint.

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertbefore"></a>InsertBefore

Facultatif. Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que `TabHome` ou `TabReview`. La valeur de l’élément est l’ID de l’onglet intégré, tel que `TabHome` ou `TabReview`.  Pour obtenir la liste des onglets intégrés, voir [OfficeTab](officetab.md). S’il est présent, il doit se trouver après **l’élément Label** . Vous ne pouvez pas avoir **à la fois InsertAfter** **et InsertBefore**.

> [!IMPORTANT]
> **L’élément InsertBefore** n’est disponible que dans PowerPoint.

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## <a name="examples"></a>Exemples

L’exemple de marques de Office ajoute le groupe de contrôles Paragraph à un onglet personnalisé et le positionnait pour qu’il apparaisse juste après un groupe personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

L’exemple de marques de Office suivant ajoute le contrôle Superscript à un groupe personnalisé et le place pour qu’il apparaisse juste après un bouton personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```
