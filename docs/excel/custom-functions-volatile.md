---
ms.date: 04/30/2019
description: Apprenez à implémenter des fonctions personnalisées de diffusion en continu et volatiles.
title: Valeurs volatiles dans les fonctions (aperçu)
localization_priority: Normal
ms.openlocfilehash: 63618adecff57398e1630e6b5ab43c0dbc753b36
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527307"
---
## <a name="volatile-values-in-functions"></a><span data-ttu-id="fd9bd-103">Valeurs volatiles dans les fonctions</span><span class="sxs-lookup"><span data-stu-id="fd9bd-103">Volatile values in functions</span></span>

<span data-ttu-id="fd9bd-104">Les fonctions volatiles sont des fonctions dans lesquelles la valeur change chaque fois que la cellule est calculée.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="fd9bd-105">La valeur peut changer même si aucun des arguments de la fonction ne change.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="fd9bd-106">Ces fonctions sont recalculées à chaque recalcul d’Excel.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="fd9bd-107">Par exemple, imaginons une cellule qui appelle la fonction `NOW`.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="fd9bd-108">Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="fd9bd-109">Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="fd9bd-110">Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="fd9bd-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="fd9bd-111">Les fonctions personnalisées vous permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="fd9bd-112">Par exemple, les [simulations Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method
) nécessitent la génération d’entrées aléatoires pour déterminer une solution optimale.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="fd9bd-113">Si vous choisissez de générer automatiquement votre fichier JSON, déclarez une fonction volatile avec la balise `@volatile`de commentaire JSDOC.</span><span class="sxs-lookup"><span data-stu-id="fd9bd-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="fd9bd-114">À partir de plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="fd9bd-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fd9bd-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fd9bd-115">See also</span></span>

* [<span data-ttu-id="fd9bd-116">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="fd9bd-116">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="fd9bd-117">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="fd9bd-117">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="fd9bd-118">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="fd9bd-118">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="fd9bd-119">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="fd9bd-119">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="fd9bd-120">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="fd9bd-120">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
