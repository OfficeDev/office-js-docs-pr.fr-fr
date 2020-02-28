---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325233"
---
# <a name="supertip"></a>Supertip

Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Titre](#title) | Oui | Texte de l’info-bulle. |
| [Description](#description) | Oui | Description de l’info-bulle.<br>**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge. |

### <a name="title"></a>Titre

Obligatoire. Texte de la propriété SuperTip. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .

### <a name="description"></a>Description

Obligatoire. Description de l’info-bulle. L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **LongStrings** de l’élément [Resources](resources.md) .

> [!NOTE]
> Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .

## <a name="example"></a>Exemple

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
