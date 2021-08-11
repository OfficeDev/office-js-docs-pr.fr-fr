---
title: Élément Script dans le fichier manifeste
description: L’élément Script définit les paramètres de script qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51902864081e135faed778de1bc6fdee15d67490de8eabc9febf493cb0c09889
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095042"
---
# <a name="script-element"></a>Élément Script

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|Éléments  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.|

## <a name="example"></a>Exemple

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
