---
ms.date: 01/14/2020
description: Découvrez comment implémenter des fonctions personnalisées de diffusion en continu volatiles et hors connexion.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: f441ef4fb7f90add5318546e3ccf4cc8bc60a8cf
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936675"
---
# <a name="volatile-values-in-functions"></a>Valeurs volatiles dans les fonctions

Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée. La valeur peut changer même si aucun des arguments de la fonction ne change. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, ce qui peut être utile lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les [simulations montes nécessitent](https://en.wikipedia.org/wiki/Monte_Carlo_method) la génération d’entrées aléatoires pour déterminer une solution optimale.

Si vous choisissez d’autogénérer votre fichier JSON, déclarez une fonction volatile avec la balise de commentaire JSDoc `@volatile` . Pour plus d’informations sur la génération automatique, voir [autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

Voici un exemple de fonction personnalisée volatile qui simule le déploiement d’un dés à six côtés.

![GIF montrant une fonction personnalisée renvoyant une valeur aléatoire pour simuler le déploiement d’un dés à six côtés.](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>Prochaines étapes
* Découvrez les [options des paramètres de fonctions personnalisées.](custom-functions-parameter-options.md)

## <a name="see-also"></a>Voir aussi

* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
