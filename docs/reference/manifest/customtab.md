---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 08/13/2021
localization_priority: Normal
ms.openlocfilehash: 3656f68a722e5e0c224f18f80a0e0214fce47cfb
ms.sourcegitcommit: bc6203dd8f21d1c375039c5ee8f1388ede9be93b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/18/2021
ms.locfileid: "58382962"
---
# <a name="customtab-element"></a>Élément CustomTab

Dans le ruban, spécifiez l’onglet et le groupe pour vos commandes de module de recherche. Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.

Sur les onglets personnalisés, le add-in peut avoir des groupes personnalisés ou intégrés. Les compléments sont limités à un onglet personnalisé.

**L’attribut id** doit être unique dans le manifeste.

> [!IMPORTANT]
> Dans Outlook Mac, l’élément n’est pas disponible, vous devez `CustomTab` donc utiliser [OfficeTab](officetab.md) à la place.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Non |  Définit un groupe de commandes.  |
|  [OfficeGroup](#officegroup)      | Non |  Représente un groupe de contrôle Office intégré. **Important**: non disponible dans Outlook. |
|  [Label](#label-tab)      | Oui |  Étiquette pour CustomTab ou Group.  |
|  [InsertAfter](#insertafter)      | Non |  Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office spécifié. **Important**: disponible uniquement dans PowerPoint. |
|  [InsertBefore](#insertbefore)      | Non |  Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office spécifié. **Important**: disponible uniquement dans PowerPoint. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si l’onglet personnalisé doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. **Important**: non disponible dans Outlook. |

### <a name="group"></a>Group

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins **un élément OfficeGroup.** Voir [Élément group](group.md). L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label.**

### <a name="officegroup"></a>OfficeGroup

Facultatif, mais s’il n’est pas présent, il doit y avoir au moins un **élément Group.** Représente un groupe de contrôle Office intégré. **L’attribut id** spécifie l’ID du groupe Office intégré. Pour trouver l’ID d’un groupe intégré, voir Rechercher les ID des contrôles et des [groupes de contrôles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) L’ordre de **groupe** et **d’OfficeGroup** dans le manifeste doit être l’ordre dans le cas où vous souhaitez qu’ils apparaissent sous l’onglet personnalisé. Ils peuvent être entremêlés s’il existe plusieurs éléments, mais tous doivent se trouver au-dessus de **l’élément Label.**

> [!IMPORTANT]
> `OfficeGroup`L’élément n’est pas disponible dans Outlook.

### <a name="label-tab"></a>Label (Tab)

Obligatoire. Étiquette de l’onglet personnalisé. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Facultatif. Spécifie que l’onglet personnalisé doit être immédiatement après un onglet Office spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ». (Voir [Rechercher les ID des contrôles et des groupes de contrôles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) S’il est présent, il doit se trouver après **l’élément Label.** Vous ne pouvez pas avoir **à la fois InsertAfter** **et InsertBefore**.

> [!IMPORTANT]
> `InsertAfter`L’élément est disponible uniquement dans PowerPoint.

### <a name="insertbefore"></a>InsertBefore

Facultatif. Spécifie que l’onglet personnalisé doit être immédiatement avant un onglet Office spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ». (Voir [Rechercher les ID des contrôles et des groupes de contrôles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  S’il est présent, il doit se trouver après **l’élément Label.** Vous ne pouvez pas avoir **à la fois InsertAfter** **et InsertBefore**.

> [!IMPORTANT]
> `InsertBefore`L’élément est disponible uniquement dans PowerPoint.

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Facultatif (booléen). Spécifie si **CustomTab** sera masqué sur les combinaisons d’applications et de plateformes qui la prise en charge d’une API qui installe un onglet contextuel personnalisé sur le ruban lors de l’utilisation. La valeur par défaut, si elle n’est pas présente, est `false` . S’il **est utilisé, OverriddenByRibbonApi doit** être le *premier* enfant de **CustomTab**. Pour plus d’informations, [voir OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!IMPORTANT]
> `OverriddenByRibbonApi`L’élément n’est pas disponible dans Outlook.

## <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
