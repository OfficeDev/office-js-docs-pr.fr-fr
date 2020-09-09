---
title: Valeurs vides et null dans les compléments Excel
description: Découvrez comment travailler avec des valeurs NULL dans les propriétés et les méthodes de modèle d’objet Excel.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409392"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a><span data-ttu-id="c4e26-103">Valeurs vides et null dans les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="c4e26-103">Blank and null values in Excel add-ins</span></span>

<span data-ttu-id="c4e26-104">`null` et les chaînes vides ont des implications particulières dans les API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="c4e26-104">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="c4e26-105">Elles sont utilisées pour représenter les cellules vides, l’absence de mise en forme ou les valeurs par défaut.</span><span class="sxs-lookup"><span data-stu-id="c4e26-105">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="c4e26-106">Cette section décrit l’utilisation de `null` et d’une chaîne vide lors de l’obtention et de la définition de propriétés.</span><span class="sxs-lookup"><span data-stu-id="c4e26-106">This section details the use of `null` and empty string when getting and setting properties.</span></span>

## <a name="null-input-in-2-d-array"></a><span data-ttu-id="c4e26-107">entrée de valeurs null dans un tableau 2D</span><span class="sxs-lookup"><span data-stu-id="c4e26-107">null input in 2-D Array</span></span>

<span data-ttu-id="c4e26-p102">Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième. Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="c4e26-p102">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="c4e26-p103">Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules. L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="c4e26-p103">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a><span data-ttu-id="c4e26-112">Entrée null pour une propriété</span><span class="sxs-lookup"><span data-stu-id="c4e26-112">null input for a property</span></span>

<span data-ttu-id="c4e26-p104">`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété `values` de la plage ne peut pas être définie sur `null`.</span><span class="sxs-lookup"><span data-stu-id="c4e26-p104">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null; // This is not a valid snippet. 
```

<span data-ttu-id="c4e26-115">De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété `color`.</span><span class="sxs-lookup"><span data-stu-id="c4e26-115">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a><span data-ttu-id="c4e26-116">valeurs de la propriété Null dans la réponse</span><span class="sxs-lookup"><span data-stu-id="c4e26-116">null property values in the response</span></span>

<span data-ttu-id="c4e26-p105">Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée. Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="c4e26-p105">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="c4e26-119">Si toutes les cellules de la plage ont la même couleur de police, `range.format.font.color` spécifie cette couleur.</span><span class="sxs-lookup"><span data-stu-id="c4e26-119">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="c4e26-120">Si plusieurs couleurs de police sont présentes dans la plage, `range.format.font.color` est `null`.</span><span class="sxs-lookup"><span data-stu-id="c4e26-120">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

## <a name="blank-input-for-a-property"></a><span data-ttu-id="c4e26-121">Entrée vide pour une propriété</span><span class="sxs-lookup"><span data-stu-id="c4e26-121">Blank input for a property</span></span>

<span data-ttu-id="c4e26-p106">Lorsque vous spécifiez une valeur vide pour une propriété (c’est-à-dire deux guillemets droits sans espace entre `''`), cela est interprété comme une instruction d’effacement ou de réinitialisation de la propriété. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="c4e26-p106">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="c4e26-124">Si vous spécifiez une valeur vide pour la propriété `values` d’une plage, le contenu de la plage est effacé.</span><span class="sxs-lookup"><span data-stu-id="c4e26-124">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="c4e26-125">Si vous spécifiez une valeur vide pour la propriété `numberFormat`, le format de nombre est réinitialisé sur `General`.</span><span class="sxs-lookup"><span data-stu-id="c4e26-125">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="c4e26-126">Si vous spécifiez une valeur vide pour les propriétés `formula` et `formulaLocale`, les valeurs de la formule sont effacées.</span><span class="sxs-lookup"><span data-stu-id="c4e26-126">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

## <a name="blank-property-values-in-the-response"></a><span data-ttu-id="c4e26-127">Valeurs de propriété vides dans la réponse</span><span class="sxs-lookup"><span data-stu-id="c4e26-127">Blank property values in the response</span></span>

<span data-ttu-id="c4e26-p107">Pour les opérations de lecture, une valeur de propriété vide dans la réponse (c'est-à-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donnée ni de valeur. Dans le premier exemple ci-dessous, la première et la dernière cellules de la plage ne contiennent pas de donnée. Dans le deuxième exemple, les deux premières cellules de la plage ne contiennent pas de formule.</span><span class="sxs-lookup"><span data-stu-id="c4e26-p107">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
