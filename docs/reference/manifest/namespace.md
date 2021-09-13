---
title: Élément Namespace dans le fichier manifest
description: L’élément Namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 3a5afed3d55bde7e9735df534215f96ae1ba7bd3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153659"
---
# <a name="namespace-element"></a>Élément Namespace

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Non  | Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md). Ne peut pas faire plus de 32 caractères. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<Namespace resid="namespace" />
```
