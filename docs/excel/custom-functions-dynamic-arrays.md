---
ms.date: 05/11/2020
description: Renvoyer plusieurs résultats à partir de votre fonction personnalisée dans un Office Excel de recherche.
title: Renvoyer plusieurs résultats à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: b7df6b2c5ca3dca24615a61e11277ac36b42c0df
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937886"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Renvoyer plusieurs résultats à partir de votre fonction personnalisée

Vous pouvez renvoyer plusieurs résultats à partir de votre fonction personnalisée qui sera renvoyée aux cellules voisines. Ce comportement est appelé débordement. Lorsque votre fonction personnalisée renvoie un tableau de résultats, elle est appelée formule de tableau dynamique. Pour plus d’informations sur les formules de tableau dynamique dans Excel, voir [Tableaux dynamiques et comportement de tableau déversé.](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)

L’image suivante montre comment la `SORT` fonction se renverse dans les cellules voisines. Votre fonction personnalisée peut également renvoyer plusieurs résultats comme celui-ci.

![Capture d’écran de la fonction « SORT » affichant plusieurs résultats dans plusieurs cellules.](../images/dynamic-array-spill.png)

Pour créer une fonction personnalisée qui est une formule de tableau dynamique, elle doit renvoyer un tableau à deux dimensions de valeurs. Si les résultats s’affichent dans des cellules voisines qui ont déjà des valeurs, la formule affiche une `#SPILL!` erreur.

L’exemple suivant montre comment renvoyer un tableau dynamique qui se déverse vers le bas.

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

L’exemple suivant montre comment renvoyer un tableau dynamique qui se déborde à droite. 

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

L’exemple suivant montre comment renvoyer un tableau dynamique qui se déborde à la fois vers le bas et la droite.

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

- [Tableaux dynamiques et comportement de tableau déversé](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Options pour Excel fonctions personnalisées](custom-functions-parameter-options.md)