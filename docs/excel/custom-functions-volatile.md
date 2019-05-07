---
ms.date: 05/03/2019
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627996"
---
# <a name="volatile-values-in-functions"></a>Valeurs volatiles dans les fonctions

Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée. La valeur peut changer même si aucun des arguments de la fonction ne change. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method
) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.

Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise `@volatile`de commentaire JSDOC. À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md).

## <a name="see-also"></a>Voir aussi

* [Options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
