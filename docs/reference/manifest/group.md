---
title: Élément Group dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7cc1f4c398eeb013eb6033b207b395466f7d72ca
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450708"
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
|  [Control](#control)    | Oui |  Ensemble d’un ou de plusieurs objets Control.  |

### <a name="label"></a>Étiquette 

Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).

### <a name="control"></a>Contrôle
Un groupe requiert au moins un contrôle.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
