---
title: Élément Page dans le fichier manifeste
description: L’élément page définit les paramètres de page HTML qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720483"
---
# <a name="page-element"></a>Élément Page

Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
