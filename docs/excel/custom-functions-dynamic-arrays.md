---
ms.date: 12/18/2019
description: Renvoyer plusieurs résultats à partir de votre fonction personnalisée dans un complément Office Excel.
title: Renvoyer plusieurs résultats à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: a2632c621071f0cbc55f545847d9e9392d884b90
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719293"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="02cc0-103">Renvoyer plusieurs résultats à partir de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="02cc0-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="02cc0-104">Vous pouvez renvoyer plusieurs résultats à partir de votre fonction personnalisée qui sera renvoyée aux cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="02cc0-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="02cc0-105">Ce comportement est appelé infiltration.</span><span class="sxs-lookup"><span data-stu-id="02cc0-105">This behavior is called spilling.</span></span> <span data-ttu-id="02cc0-106">Lorsque votre fonction personnalisée renvoie un tableau de résultats, il s’agit d’une formule matricielle dynamique.</span><span class="sxs-lookup"><span data-stu-id="02cc0-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="02cc0-107">Pour plus d’informations sur les formules de tableau dynamique dans Excel, voir [tableaux dynamiques et comportement de tableau propagé](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span><span class="sxs-lookup"><span data-stu-id="02cc0-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="02cc0-108">L’image suivante montre comment la `SORT` fonction descend en cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="02cc0-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="02cc0-109">Votre fonction personnalisée peut également renvoyer plusieurs résultats de la manière suivante.</span><span class="sxs-lookup"><span data-stu-id="02cc0-109">Your custom function can also return multiple results like this.</span></span>

![Capture d’écran de la fonction « Trier » affichant plusieurs résultats en plusieurs cellules.](../images/dynamic-array-spill.png)

<span data-ttu-id="02cc0-111">Pour créer une fonction personnalisée qui est une formule matricielle dynamique, elle doit renvoyer un tableau à deux dimensions de valeurs.</span><span class="sxs-lookup"><span data-stu-id="02cc0-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="02cc0-112">Si les résultats sont détourés en cellules voisines qui contiennent déjà des valeurs, la `#SPILL!` formule affiche une erreur.</span><span class="sxs-lookup"><span data-stu-id="02cc0-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="02cc0-113">L’exemple suivant montre comment retourner un tableau dynamique qui se renverse vers le bas.</span><span class="sxs-lookup"><span data-stu-id="02cc0-113">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="02cc0-114">L’exemple suivant montre comment renvoyer un tableau dynamique qui se remet à la droite.</span><span class="sxs-lookup"><span data-stu-id="02cc0-114">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="02cc0-115">L’exemple suivant montre comment retourner un tableau dynamique qui renverse les deux à la fois.</span><span class="sxs-lookup"><span data-stu-id="02cc0-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="02cc0-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="02cc0-116">See also</span></span>

- [<span data-ttu-id="02cc0-117">Tableaux dynamiques et comportement de tableau renversé</span><span class="sxs-lookup"><span data-stu-id="02cc0-117">Dynamic arrays and spilled array behavior</span></span>](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="02cc0-118">Options pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="02cc0-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)