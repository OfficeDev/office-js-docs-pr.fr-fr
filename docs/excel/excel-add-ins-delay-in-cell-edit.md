---
title: Retarder l’exécution pendant la modification de la cellule
description: Découvrez comment retarder l’exécution de la méthode Excel.run lorsqu’une cellule est en cours de modification.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c5609fbb2a39d6ecc69063d4bccdfbc1da1c102d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340805"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Retarder l’exécution pendant la modification de la cellule

`Excel.run`a une surcharge qui prend en charge une [Excel. Objet RunOptions](/javascript/api/excel/excel.runoptions). Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution. La propriété suivante est actuellement prise en charge.

- `delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule. Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur). Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
