---
title: Élément Supertip dans le fichier manifest
description: L’élément Supertip définit une boîte à outils enrichie (titre et description).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936186"
---
# <a name="supertip"></a>Supertip

Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Titre](#title) | Oui | Texte de l’info-bulle. |
| [Description](#description) | Oui | Description de l’info-bulle.<br>**Remarque**: (Outlook) Seuls les clients Windows mac sont pris en charge. |

### <a name="title"></a>Titre

Obligatoire. Texte de la propriété SuperTip. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources.](resources.md)

### <a name="description"></a>Description

Obligatoire. Description de l’info-bulle. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources.](resources.md)

> [!NOTE]
> Par Outlook, seuls Windows clients Mac et mac supportent **l’élément Description.**

## <a name="example"></a>Exemple

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
