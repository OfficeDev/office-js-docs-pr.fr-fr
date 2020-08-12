---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641381"
---
# <a name="sourcelocation-element-custom-functions"></a>Élément SourceLocation (fonctions personnalisées)

Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.

## <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Oui      | Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste. |

## <a name="child-elements"></a>Child, élément

Aucun

## <a name="example"></a>Exemple

```xml
<SourceLocation resid="pageURL"/>
```
