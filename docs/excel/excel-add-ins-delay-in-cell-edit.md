---
title: Retarder l’exécution pendant la modification de la cellule
description: Découvrez comment retarder l’exécution de la méthode Excel.run lorsqu’une cellule est en cours de modification.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: b7b28064ef4d313639391e63cba780351b5623f9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349517"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="8f974-103">Retarder l’exécution pendant la modification de la cellule</span><span class="sxs-lookup"><span data-stu-id="8f974-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="8f974-104">`Excel.run`a une surcharge qui prend en [charge une Excel. Objet RunOptions.](/javascript/api/excel/excel.runoptions)</span><span class="sxs-lookup"><span data-stu-id="8f974-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="8f974-105">Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="8f974-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="8f974-106">La propriété suivante est actuellement prise en charge.</span><span class="sxs-lookup"><span data-stu-id="8f974-106">The following property is currently supported.</span></span>

- <span data-ttu-id="8f974-107">`delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="8f974-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="8f974-108">Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="8f974-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="8f974-109">Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur).</span><span class="sxs-lookup"><span data-stu-id="8f974-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="8f974-110">Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.</span><span class="sxs-lookup"><span data-stu-id="8f974-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
