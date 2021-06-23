---
title: Supprimer les doublons à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour supprimer les doublons.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 859214d36bdf66a284304ba1d5f7f2d642b718cb
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075767"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a><span data-ttu-id="420fb-103">Supprimer les doublons à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="420fb-103">Remove duplicates using the Excel JavaScript API</span></span>

<span data-ttu-id="420fb-104">Cet article fournit un exemple de code qui supprime les entrées en double dans une plage à l’aide Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="420fb-104">This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API.</span></span> <span data-ttu-id="420fb-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="420fb-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="remove-rows-with-duplicate-entries"></a><span data-ttu-id="420fb-106">Supprimer des lignes avec des entrées en double</span><span class="sxs-lookup"><span data-stu-id="420fb-106">Remove rows with duplicate entries</span></span>

<span data-ttu-id="420fb-107">La [méthode Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) supprime les lignes avec des entrées en double dans les colonnes spécifiées.</span><span class="sxs-lookup"><span data-stu-id="420fb-107">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="420fb-108">La méthode passe par chaque ligne de la plage, de l’index à la valeur la plus faible à l’index à valeur la plus élevée de la plage (du haut vers le bas).</span><span class="sxs-lookup"><span data-stu-id="420fb-108">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="420fb-109">Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage.</span><span class="sxs-lookup"><span data-stu-id="420fb-109">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="420fb-110">Les rangées de la plage en-dessous de la rangée supprimée sont déplacées.</span><span class="sxs-lookup"><span data-stu-id="420fb-110">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="420fb-111">`removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.</span><span class="sxs-lookup"><span data-stu-id="420fb-111">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="420fb-112">`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons.</span><span class="sxs-lookup"><span data-stu-id="420fb-112">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="420fb-113">Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="420fb-113">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="420fb-114">La méthode prend également un paramètre booléen qui spécifie si la première ligne est un en-tête.</span><span class="sxs-lookup"><span data-stu-id="420fb-114">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="420fb-115">Lorsque **true**, la rangée du dessus est ignorée lorsque les doublons sont pris en considération.</span><span class="sxs-lookup"><span data-stu-id="420fb-115">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="420fb-116">La méthode renvoie un objet qui spécifie le nombre de lignes supprimées et le nombre de lignes `removeDuplicates` `RemoveDuplicatesResult` uniques restantes.</span><span class="sxs-lookup"><span data-stu-id="420fb-116">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="420fb-117">Lorsque vous utilisez la méthode `removeDuplicates` d’une plage, gardez les données suivantes à l’esprit :</span><span class="sxs-lookup"><span data-stu-id="420fb-117">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="420fb-118">`removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction.</span><span class="sxs-lookup"><span data-stu-id="420fb-118">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="420fb-119">Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.</span><span class="sxs-lookup"><span data-stu-id="420fb-119">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="420fb-120">Les cellules vides ne sont pas ignorées par`removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="420fb-120">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="420fb-121">La valeur d’une cellule vide est traitée comme toute autre valeur.</span><span class="sxs-lookup"><span data-stu-id="420fb-121">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="420fb-122">Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="420fb-122">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="420fb-123">L’exemple de code suivant montre la suppression des entrées avec des valeurs en double dans la première colonne.</span><span class="sxs-lookup"><span data-stu-id="420fb-123">The following code sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a><span data-ttu-id="420fb-124">Données avant la suppression des entrées en double</span><span class="sxs-lookup"><span data-stu-id="420fb-124">Data before duplicate entries are removed</span></span>

![Données dans Excel méthode de suppression des doublons de la plage a été exécuté.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a><span data-ttu-id="420fb-126">Données après suppression des entrées en double</span><span class="sxs-lookup"><span data-stu-id="420fb-126">Data after duplicate entries are removed</span></span>

![Données dans Excel une fois que la méthode des doublons de suppression de plage a été exécuté.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="420fb-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="420fb-128">See also</span></span>

- [<span data-ttu-id="420fb-129">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="420fb-129">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="420fb-130">Utiliser des cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="420fb-130">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="420fb-131">Couper, copier et coller des plages à l’aide de l Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="420fb-131">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-cut-copy-paste.md)
- [<span data-ttu-id="420fb-132">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="420fb-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
