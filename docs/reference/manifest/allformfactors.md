---
title: Élément AllFormFactors dans le fichier manifeste
description: Spécifie les paramètres d’un complément pour tous les facteurs de forme.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 674fbe9defa961cb0eef1103cf2dedea0983ffabadc665b172d1f3b15292e987
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088538"
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
