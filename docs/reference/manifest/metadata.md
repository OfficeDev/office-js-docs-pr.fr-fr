---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6f58b00bb13bde1e2b1742462716119b8b6d369d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152865"
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
