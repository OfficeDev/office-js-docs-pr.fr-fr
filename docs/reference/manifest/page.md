---
title: Élément Page dans le fichier manifeste
description: L’élément Page définit les paramètres de page HTML qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937841"
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
