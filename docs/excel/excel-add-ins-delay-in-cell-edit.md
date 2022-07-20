---
title: Retarder l’exécution pendant la modification de la cellule
description: Découvrez comment retarder l’exécution de la méthode Excel.run lors de la modification d’une cellule.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1abcdb382150db486033b32d2521207ab0b7f28f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889218"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Retarder l’exécution pendant la modification de la cellule

`Excel.run` a une surcharge qui accepte un objet [Excel.RunOptions](/javascript/api/excel/excel.runoptions) . Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge.

- `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Quand `true`, la demande de lot est retardée et s’exécute lorsque l’utilisateur quitte le mode d’édition de la cellule. Quand `false`, la demande de lot échoue automatiquement si l’utilisateur est en mode d’édition de cellule (ce qui provoque une erreur pour atteindre l’utilisateur). Le comportement par défaut sans `delayForCellEdit` propriété spécifiée est équivalent au moment où il est `false`.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
