---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450687"
---
# <a name="sourcelocation-element"></a>SourceLocation, élément

Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.

## <a name="attributes"></a>Attributs

| **Attribut** | **Obligatoire** | **Description**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Oui          | Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste. |

## <a name="child-elements"></a>Child, élément

Aucun

## <a name="example"></a>Exemple

```xml
<SourceLocation resid="pageURL"/>
```
