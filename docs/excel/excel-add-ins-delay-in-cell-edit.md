---
title: Retarder l’exécution pendant que la cellule est en cours de modification
description: Découvrez comment retarder l’exécution de la méthode Excel. Run quand une cellule est en cours de modification.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: eb33f4cb7cce3b1f8642e00f432e708e90b5b895
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409389"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Retarder l’exécution pendant que la cellule est en cours de modification

`Excel.run` possède une surcharge qui prend un objet [Excel. RunOptions](/javascript/api/excel/excel.runoptions) . Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge :

* `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
