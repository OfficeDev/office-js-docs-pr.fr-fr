---
title: Valeurs vides et null dans les compléments Excel
description: Découvrez comment travailler avec des valeurs nulles vides dans Excel méthodes et propriétés du modèle objet.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937679"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a>Valeurs vides et null dans les compléments Excel

`null` et les chaînes vides ont des implications particulières dans les API JavaScript Excel. Elles sont utilisées pour représenter les cellules vides, l’absence de mise en forme ou les valeurs par défaut. Cette section décrit l’utilisation de `null` et d’une chaîne vide lors de l’obtention et de la définition de propriétés.

## <a name="null-input-in-2-d-array"></a>entrée de valeurs null dans un tableau 2D

Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième. Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.

Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules. L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a>Entrée null pour une propriété

`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété `values` de la plage ne peut pas être définie sur `null`.

```js
range.values = null; // This is not a valid snippet. 
```

De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété `color`.

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a>valeurs de la propriété Null dans la réponse

Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée. Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color`:

* Si toutes les cellules de la plage ont la même couleur de police, `range.format.font.color` spécifie cette couleur.
* Si plusieurs couleurs de police sont présentes dans la plage, `range.format.font.color` est `null`.

## <a name="blank-input-for-a-property"></a>Entrée vide pour une propriété

Lorsque vous spécifiez une valeur vide pour une propriété (c’est-à-dire deux guillemets droits sans espace entre `''`), cela est interprété comme une instruction d’effacement ou de réinitialisation de la propriété. Par exemple :

* Si vous spécifiez une valeur vide pour la propriété `values` d’une plage, le contenu de la plage est effacé.
* Si vous spécifiez une valeur vide pour la propriété `numberFormat`, le format de nombre est réinitialisé sur `General`.
* Si vous spécifiez une valeur vide pour les propriétés `formula` et `formulaLocale`, les valeurs de la formule sont effacées.

## <a name="blank-property-values-in-the-response"></a>Valeurs de propriété vides dans la réponse

Pour les opérations de lecture, une valeur de propriété vide dans la réponse (c'est-à-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donnée ni de valeur. Dans le premier exemple ci-dessous, la première et la dernière cellules de la plage ne contiennent pas de donnée. Dans le deuxième exemple, les deux premières cellules de la plage ne contiennent pas de formule.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
