---
title: Élément Namespace dans le fichier manifest
description: L’élément namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771257"
---
# <a name="namespace-element"></a>Élément Namespace

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Non  | Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md). Il ne peut pas comporter plus de 32 caractères. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<Namespace resid="namespace" />
```
