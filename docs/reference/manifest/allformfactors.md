---
title: Élément AllFormFactors dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450736"
---
# <a name="allformfactors-element"></a>AllFormFactors, élément

Spécifie les paramètres d’un complément pour tous les facteurs de forme. Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées. **AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
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
