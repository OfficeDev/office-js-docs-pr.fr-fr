---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718068"
---
# <a name="metadata-element"></a>Élément de métadonnées

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
