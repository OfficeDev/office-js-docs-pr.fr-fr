---
ms.date: 12/18/2019
description: Renvoyer plusieurs résultats à partir de votre fonction personnalisée dans un complément Office Excel.
title: Renvoyer plusieurs résultats à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: 753755b481ab3db0de711c80ef082aedc82177ae
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217836"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Renvoyer plusieurs résultats à partir de votre fonction personnalisée

Vous pouvez renvoyer plusieurs résultats à partir de votre fonction personnalisée qui sera renvoyée aux cellules voisines. Ce comportement est appelé infiltration. Lorsque votre fonction personnalisée renvoie un tableau de résultats, il s’agit d’une formule matricielle dynamique. Pour plus d’informations sur les formules de tableau dynamique dans Excel, voir [tableaux dynamiques et comportement de tableau propagé](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

L’image suivante montre comment la `SORT` fonction descend en cellules voisines. Votre fonction personnalisée peut également renvoyer plusieurs résultats de la manière suivante.

![Capture d’écran de la fonction « Trier » affichant plusieurs résultats en plusieurs cellules.](../images/dynamic-array-spill.png)

Pour créer une fonction personnalisée qui est une formule matricielle dynamique, elle doit renvoyer un tableau à deux dimensions de valeurs. Si les résultats sont détourés en cellules voisines qui contiennent déjà des valeurs, la formule affiche une `#SPILL!` erreur.

L’exemple suivant montre comment retourner un tableau dynamique qui se renverse vers le bas.

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

L’exemple suivant montre comment renvoyer un tableau dynamique qui se remet à la droite. 

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

L’exemple suivant montre comment retourner un tableau dynamique qui renverse les deux à la fois.

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

## <a name="see-also"></a>Voir aussi

- [Tableaux dynamiques et comportement de tableau renversé](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Options pour les fonctions personnalisées Excel](custom-functions-parameter-options.md)