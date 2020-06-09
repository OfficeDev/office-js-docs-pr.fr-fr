---
ms.date: 01/14/2020
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: a318c87cc5b5f45bf3b1f5fe1341b7008f5a3d2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609789"
---
# <a name="volatile-values-in-functions"></a>Valeurs volatiles dans les fonctions

Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée. La valeur peut changer même si aucun des arguments de la fonction ne change. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.

Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise de commentaire JSDoc `@volatile` . À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

Voici un exemple de fonction personnalisée volatile qui simule le roulement d’un dés à six côtés.

![Image gif illustrant une fonction personnalisée renvoyant une valeur aléatoire pour simuler le roulement d’un dés à six côtés](../images/six-sided-die.gif)

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

## <a name="next-steps"></a>Étapes suivantes
* En savoir plus sur les [Options des paramètres des fonctions personnalisées](custom-functions-parameter-options.md).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
