---
ms.date: 07/15/2019
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: 92d61aff4c3f4b4cbc79a3981db12ed1ce0ffb9d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771645"
---
# <a name="volatile-values-in-functions"></a>Valeurs volatiles dans les fonctions

Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée. La valeur peut changer même si aucun des arguments de la fonction ne change. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.

Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise `@volatile`de commentaire JSDoc. À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

Voici un exemple de fonction personnalisée volatile qui simule le roulement d’un dés à six côtés.

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md).

## <a name="see-also"></a>Voir aussi

* [Options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
