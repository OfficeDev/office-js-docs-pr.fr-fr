---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f2d881f31f4e46e7f5bb8ab30d78abd0e9b7200
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990683"
---
# <a name="sourcelocation-element-custom-functions"></a>Élément SourceLocation (fonctions personnalisées)

Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.

**Type de add-in :** Fonction personnalisée

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
