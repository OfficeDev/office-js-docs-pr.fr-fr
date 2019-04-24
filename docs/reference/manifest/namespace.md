---
title: Élément Namespace dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452101"
---
# <a name="namespace-element"></a>Élément Namespace

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Oui  | Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md). |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<Namespace resid="namespace" />
```
