---
title: Élément SourceLocation dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612311"
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
