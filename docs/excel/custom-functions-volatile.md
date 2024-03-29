---
ms.date: 01/14/2020
description: Découvrez comment implémenter des fonctions personnalisées de diffusion en continu volatiles et hors connexion.
title: Valeurs volatiles dans les fonctions
ms.localizationpriority: medium
ms.openlocfilehash: 401be3e04a7b36a226547175df4311fc653c027a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744469"
---
# <a name="volatile-values-in-functions"></a>Valeurs volatiles dans les fonctions

Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée. La valeur peut changer même si aucun des arguments de la fonction ne change. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, ce qui peut être utile lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les [simulations Monte Contrôle nécessitent](https://en.wikipedia.org/wiki/Monte_Carlo_method) la génération d’entrées aléatoires pour déterminer une solution optimale.

Si vous choisissez d’autogénérer votre fichier JSON, déclarez une fonction volatile avec la balise de commentaire `@volatile`JSDoc . Pour plus d’informations sur la génération automatique, voir [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

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
* Découvrez les [options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md).

## <a name="see-also"></a>Voir aussi

* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
