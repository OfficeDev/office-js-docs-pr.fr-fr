---
title: Retarder l’exécution pendant la modification de la cellule
description: Découvrez comment retarder l’exécution de la méthode Excel.run lorsqu’une cellule est en cours de modification.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 41bbfba3894bcef0c1fd075ce76557dfdc4ba4721b7bc7b19ca21756b86ccc4d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084282"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Retarder l’exécution pendant la modification de la cellule

`Excel.run`a une surcharge qui prend en [charge une Excel. Objet RunOptions.](/javascript/api/excel/excel.runoptions) Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge.

- `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
