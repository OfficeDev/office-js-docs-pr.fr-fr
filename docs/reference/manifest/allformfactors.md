---
title: Élément AllFormFactors dans le fichier manifeste
description: Spécifie les paramètres d’un complément pour tous les facteurs de forme.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608795"
---
# <a name="allformfactors-element"></a>AllFormFactors, élément

Spécifie les paramètres d’un complément pour tous les facteurs de forme. Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées. **AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Oui |  Définit l’emplacement où se trouvent les fonctionnalités d’un complément |

## <a name="allformfactors-example"></a>Exemple AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
