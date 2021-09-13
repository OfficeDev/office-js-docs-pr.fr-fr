---
title: Retarder l’exécution pendant la modification de la cellule
description: Découvrez comment retarder l’exécution de la méthode Excel.run lorsqu’une cellule est en cours de modification.
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 246faebf593e16b342606d975573a4c29279cc42
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150495"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Retarder l’exécution pendant la modification de la cellule

`Excel.run`a une surcharge qui prend en [charge une Excel. Objet RunOptions.](/javascript/api/excel/excel.runoptions) Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge.

- `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
