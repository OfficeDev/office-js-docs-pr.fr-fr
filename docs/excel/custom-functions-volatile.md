---
ms.date: 06/17/2019
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions
localization_priority: Normal
ms.openlocfilehash: 0edf4071ce366c40300663233f1de318a544169b
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059845"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="c564e-103">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="c564e-103">Volatile values in functions</span></span>

<span data-ttu-id="c564e-104">Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée.</span><span class="sxs-lookup"><span data-stu-id="c564e-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="c564e-105">La valeur peut changer même si aucun des arguments de la fonction ne change.</span><span class="sxs-lookup"><span data-stu-id="c564e-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="c564e-106">Ces fonctions sont recalculées à chaque recalcul d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c564e-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="c564e-107">Par exemple, imaginons une cellule qui appelle la fonction `NOW`.</span><span class="sxs-lookup"><span data-stu-id="c564e-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="c564e-108">Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.</span><span class="sxs-lookup"><span data-stu-id="c564e-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="c564e-109">Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="c564e-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="c564e-110">Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="c564e-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="c564e-111">Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation.</span><span class="sxs-lookup"><span data-stu-id="c564e-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="c564e-112">Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.</span><span class="sxs-lookup"><span data-stu-id="c564e-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="c564e-113">Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise `@volatile`de commentaire JSDOC.</span><span class="sxs-lookup"><span data-stu-id="c564e-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="c564e-114">À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="c564e-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="c564e-115">Voici un exemple de fonction personnalisée volatile qui simule le roulement d’un dés à six côtés.</span><span class="sxs-lookup"><span data-stu-id="c564e-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="c564e-116">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="c564e-116">Next steps</span></span>
<span data-ttu-id="c564e-117">Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md).</span><span class="sxs-lookup"><span data-stu-id="c564e-117">Learn how to [save state in your custom functions](custom-functions-save-state.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c564e-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c564e-118">See also</span></span>

* [<span data-ttu-id="c564e-119">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c564e-119">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="c564e-120">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c564e-120">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c564e-121">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="c564e-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
