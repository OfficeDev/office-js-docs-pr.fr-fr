---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938101"
---
# <a name="metadata-element"></a>Élément de métadonnées

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
