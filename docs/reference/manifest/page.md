---
title: Élément Page dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452073"
---
# <a name="page-element"></a>Élément Page

Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
