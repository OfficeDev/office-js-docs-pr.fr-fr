---
title: Élément Group dans le fichier manifeste
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: ad1a566e259188ed20032bc5a3004736474e1f01
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670131"
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
Un groupe requiert au moins un contrôle. Pour plus d’informations sur les types de contrôles pris en charge, reportez-vous à l’élément [Control](control.md) .

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
