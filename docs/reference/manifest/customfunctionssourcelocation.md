---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 08/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: 84d5607fbb02c1925137e1a143b7715c7c87c6fa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150604"
---
# <a name="sourcelocation-element-custom-functions"></a>Élément SourceLocation (fonctions personnalisées)

Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.

## <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Oui      | Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste. Ne peut pas faire plus de 32 caractères. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<SourceLocation resid="pageURL"/>
```
