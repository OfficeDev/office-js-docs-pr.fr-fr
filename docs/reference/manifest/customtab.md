---
title: Élément CustomTab dans le fichier manifest
description: Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771325"
---
# <a name="customtab-element"></a>Élément CustomTab

Dans le ruban, spécifiez l’onglet et le groupe pour vos commandes de complément. Il peut s’agir de l’onglet par défaut (**Accueil**, **Message** ou **Réunion**) ou un onglet personnalisé défini par le complément.

Dans les onglets personnalisés, le complément peut contenir des groupes personnalisés ou intégrés. Les compléments sont limités à un onglet personnalisé.

L’attribut **ID** doit être unique dans le manifeste.

> [!IMPORTANT]
> Dans Outlook sur Mac, l' `CustomTab` élément n’est pas disponible et vous devez utiliser [OfficeTab](officetab.md) à la place.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Non |  Définit un groupe de commandes.  |
|  [OfficeGroup](#officegroup)      | Non |  Représente un groupe de contrôles Office prédéfini.  |
|  [Label](#label-tab)      | Oui |  Étiquette pour CustomTab ou Group.  |
|  [InsertAfter](#insertafter)      | Non |  Spécifie que l’onglet personnalisé doit se trouver immédiatement après un onglet Office prédéfini spécifié.  |
|  [InsertBefore](#insertbefore)      | Non |  Spécifie que l’onglet personnalisé doit se trouver immédiatement avant un onglet Office prédéfini spécifié.  |

### <a name="group"></a>Groupe

Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un élément **OfficeGroup** . Voir [Élément group](group.md). L’ordre des **groupes** et des **OfficeGroup** dans le manifeste doit être l’ordre dans lequel vous souhaitez qu’ils apparaissent dans l’onglet personnalisé. Ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être au-dessus de l’élément **label** .

### <a name="officegroup"></a>OfficeGroup

Facultatif, mais si ce n’est pas le cas, il doit y avoir au moins un élément **Group** . Représente un groupe de contrôles Office prédéfini. L’attribut **ID** spécifie l’ID du groupe Office prédéfini. Pour Rechercher l’ID d’un groupe prédéfini, voir [Rechercher les ID de contrôles et les groupes](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)de contrôles. L’ordre des **groupes** et des **OfficeGroup** dans le manifeste doit être l’ordre dans lequel vous souhaitez qu’ils apparaissent dans l’onglet personnalisé. Ils peuvent être mélangés s’il y a plusieurs éléments, mais ils doivent tous être au-dessus de l’élément **label** .

### <a name="label-tab"></a>Label (Tab)

Obligatoire. Étiquette de l’onglet personnalisé. L’attribut **RESID** ne peut pas contenir plus de 32 caractères et doit être défini sur la valeur de l’attribut **ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .

### <a name="insertafter"></a>InsertAfter

Facultatif. Spécifie que l’onglet personnalisé doit se trouver immédiatement après un onglet Office prédéfini spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ». (Voir [Rechercher les ID des contrôles et des groupes de](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)contrôles.) Le cas échéant, doit se trouver après l’élément **label** . Vous ne pouvez pas avoir à la fois **InsertAfter** et **InsertBefore**.

### <a name="insertbefore"></a>InsertBefore

Facultatif. Spécifie que l’onglet personnalisé doit se trouver immédiatement avant un onglet Office prédéfini spécifié. La valeur de l’élément est l’ID de l’onglet intégré, tel que « TabHome » ou « TabReview ». (Voir [Rechercher les ID des contrôles et des groupes de](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)contrôles.)  Le cas échéant, doit se trouver après l’élément **label** . Vous ne pouvez pas avoir à la fois **InsertAfter** et **InsertBefore**.

## <a name="customtab-example"></a>Exemple CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
