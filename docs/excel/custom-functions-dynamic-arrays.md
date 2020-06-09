---
ms.date: 05/11/2020
description: Renvoyer plusieurs résultats à partir de votre fonction personnalisée dans un complément Office Excel.
title: Renvoyer plusieurs résultats à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: e25965277fbbe1c39007f79f401bf62b25760488
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609649"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="db030-103">Renvoyer plusieurs résultats à partir de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="db030-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="db030-104">Vous pouvez renvoyer plusieurs résultats à partir de votre fonction personnalisée qui sera renvoyée aux cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="db030-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="db030-105">Ce comportement est appelé infiltration.</span><span class="sxs-lookup"><span data-stu-id="db030-105">This behavior is called spilling.</span></span> <span data-ttu-id="db030-106">Lorsque votre fonction personnalisée renvoie un tableau de résultats, il s’agit d’une formule matricielle dynamique.</span><span class="sxs-lookup"><span data-stu-id="db030-106">When your custom function returns an array of results, it's known as a dynamic array formula.</span></span> <span data-ttu-id="db030-107">Pour plus d’informations sur les formules de tableau dynamique dans Excel, voir [tableaux dynamiques et comportement de tableau propagé](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span><span class="sxs-lookup"><span data-stu-id="db030-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="db030-108">L’image suivante montre comment la `SORT` fonction descend en cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="db030-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="db030-109">Votre fonction personnalisée peut également renvoyer plusieurs résultats de la manière suivante.</span><span class="sxs-lookup"><span data-stu-id="db030-109">Your custom function can also return multiple results like this.</span></span>

![Capture d’écran de la fonction « Trier » affichant plusieurs résultats en plusieurs cellules.](../images/dynamic-array-spill.png)

<span data-ttu-id="db030-111">Pour créer une fonction personnalisée qui est une formule matricielle dynamique, elle doit renvoyer un tableau à deux dimensions de valeurs.</span><span class="sxs-lookup"><span data-stu-id="db030-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="db030-112">Si les résultats sont détourés en cellules voisines qui contiennent déjà des valeurs, la formule affiche une `#SPILL!` erreur.</span><span class="sxs-lookup"><span data-stu-id="db030-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="db030-113">L’exemple suivant montre comment retourner un tableau dynamique qui se renverse vers le bas.</span><span class="sxs-lookup"><span data-stu-id="db030-113">The following example shows how to return a dynamic array that spills down.</span></span>

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

<span data-ttu-id="db030-114">L’exemple suivant montre comment renvoyer un tableau dynamique qui se remet à la droite.</span><span class="sxs-lookup"><span data-stu-id="db030-114">The following example shows how to return a dynamic array that spills right.</span></span> 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

<span data-ttu-id="db030-115">L’exemple suivant montre comment retourner un tableau dynamique qui renverse les deux à la fois.</span><span class="sxs-lookup"><span data-stu-id="db030-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a><span data-ttu-id="db030-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="db030-116">See also</span></span>

- [<span data-ttu-id="db030-117">Tableaux dynamiques et comportement de tableau renversé</span><span class="sxs-lookup"><span data-stu-id="db030-117">Dynamic arrays and spilled array behavior</span></span>](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="db030-118">Options pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="db030-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)