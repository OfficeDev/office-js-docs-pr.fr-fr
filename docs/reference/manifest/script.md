---
title: Élément Script dans le fichier manifeste
description: L’élément Script définit les paramètres de script qu’une fonction personnalisée utilise dans Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 259976f752cf3fca72c5012bedd92b9bf021f6aa
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990669"
---
# <a name="script-element"></a>Élément Script

Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.

**Type de add-in :** Fonction personnalisée

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
