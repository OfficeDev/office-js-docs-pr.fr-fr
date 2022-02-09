---
title: Élément Control dans le fichier manifeste
description: Définit un contrôle qui exécute une action ou lance un volet Des tâches.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa7ff9b0162070b378352ce187de15a34323b998
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467835"
---
# <a name="control-element"></a>Élément Control

Définit un contrôle qui exécute une action ou lance un volet Des tâches. Un élément **Control** peut être une option de bouton ou de menu. Au moins un élément **Control** doit être inclus dans un élément [Group](group.md).

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Pour un add-in de volet de tâches.)
- Certains éléments enfants peuvent être associés à des ensembles de conditions requises supplémentaires.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|**xsi:type**|Oui|Type de contrôle défini. Peut être `Button`, `Menu`ou `MobileButton`. |
|**id**|Oui|ID de l’élément de contrôle. Il doit comporter 125 caractères au maximum. Doit être unique dans tous **les éléments Control** du manifeste.|

> [!NOTE]
> La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Elle s’applique uniquement aux éléments **Control** contenus dans un élément [MobileFormFactor](mobileformfactor.md).

## <a name="child-elements"></a>Éléments enfants

Les éléments enfants valides dépendent de la valeur de l’attribut **xsi:type** .

- [Type de bouton de l’élément Control](control-button.md)
- [Type de menu de l’élément Control](control-menu.md)
- [Type MobileButton de l’élément Control](control-mobilebutton.md)
