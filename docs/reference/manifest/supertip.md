---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659654"
---
# <a name="supertip"></a>Supertip

Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Titre](#title) | Oui | Texte de l’info-bulle. |
| [Description](#description) | Oui | Description de l’info-bulle.<br>**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge. |

### <a name="title"></a>Titre

Obligatoire. Texte de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).

### <a name="description"></a>Description

Obligatoire. Description de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).

> [!NOTE]
> Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .

## <a name="example"></a>Exemple

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
