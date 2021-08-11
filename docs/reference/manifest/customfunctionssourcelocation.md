---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: b18a340d4dd4403b1e5fd2c7d8868a820eef5a241ac3d666926d8f2cb49fcc09
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098300"
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
