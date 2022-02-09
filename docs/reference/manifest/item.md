---
title: Élément Item dans le fichier manifeste
description: Spécifie un élément dans un menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: cd46b46e1466b8cb9bab7e283ddca437721e762e
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467890"
---
# <a name="item-element"></a>Élément Item

Spécifie un élément dans un menu.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Oui |  Texte du bouton. |
|  [Supertip](supertip.md)  | Oui |  Info-bulle pour le bouton.    |
|  [Icon](icon.md)      | Oui |  Image du bouton.         |
|  [Action](action.md)    | Oui |  Spécifie l’action à effectuer. Il ne peut y avoir **qu’un seul enfant Action** **d’un élément Item** .  |
|  [Enabled](enabled.md)    | Non |  Spécifie si le contrôle est activé au lancement du module.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le bouton doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. S’il est utilisé, il doit s’agit du *premier* élément enfant. |

### <a name="label"></a>Étiquette

Spécifie le texte du bouton au moyen de son seul attribut, **resid**, qui ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’enfant **ShortStrings** de l’élément [Resources](resources.md) .

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="examples"></a>Exemples

Pour obtenir des exemples, [voir Contrôle de type Menu](control-menu.md).