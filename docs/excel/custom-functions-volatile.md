---
ms.date: 01/14/2020
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: 7545d9928eaeb3779a8f7e04c87d0d5f33a7a131
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275776"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="f5510-103">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="f5510-103">Volatile values in functions</span></span>

<span data-ttu-id="f5510-104">Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée.</span><span class="sxs-lookup"><span data-stu-id="f5510-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="f5510-105">La valeur peut changer même si aucun des arguments de la fonction ne change.</span><span class="sxs-lookup"><span data-stu-id="f5510-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="f5510-106">Ces fonctions sont recalculées à chaque recalcul d’Excel.</span><span class="sxs-lookup"><span data-stu-id="f5510-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="f5510-107">Par exemple, imaginons une cellule qui appelle la fonction `NOW`.</span><span class="sxs-lookup"><span data-stu-id="f5510-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="f5510-108">Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.</span><span class="sxs-lookup"><span data-stu-id="f5510-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f5510-109">Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="f5510-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="f5510-110">Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="f5510-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="f5510-111">Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation.</span><span class="sxs-lookup"><span data-stu-id="f5510-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="f5510-112">Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.</span><span class="sxs-lookup"><span data-stu-id="f5510-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="f5510-113">Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise de commentaire JSDoc `@volatile` .</span><span class="sxs-lookup"><span data-stu-id="f5510-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="f5510-114">À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="f5510-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="f5510-115">Voici un exemple de fonction personnalisée volatile qui simule le roulement d’un dés à six côtés.</span><span class="sxs-lookup"><span data-stu-id="f5510-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="f5510-117">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f5510-117">Next steps</span></span>
* <span data-ttu-id="f5510-118">En savoir plus sur les [Options des paramètres des fonctions personnalisées](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="f5510-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f5510-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f5510-119">See also</span></span>

* [<span data-ttu-id="f5510-120">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f5510-120">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f5510-121">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="f5510-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
