---
title: Couper, copier et coller des plages à l’aide de l Excel API JavaScript
description: Découvrez comment couper, copier et coller des plages à l’aide de l Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2112702110b72e0020ed72090ce495abb3ff5366
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075823"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="03314-103">Couper, copier et coller des plages à l’aide de l Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="03314-103">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="03314-104">Cet article fournit des exemples de code qui coupent, copient et collent des plages à l’aide Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="03314-104">This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="03314-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="03314-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a><span data-ttu-id="03314-106">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="03314-106">Copy and paste</span></span>

<span data-ttu-id="03314-107">La [méthode Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) réplique les **actions** **Copier** et coller de l’interface Excel’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="03314-107">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="03314-108">La destination est `Range` l’objet `copyFrom` qui est appelé.</span><span class="sxs-lookup"><span data-stu-id="03314-108">The destination is the `Range` object that `copyFrom` is called on.</span></span> <span data-ttu-id="03314-109">La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.</span><span class="sxs-lookup"><span data-stu-id="03314-109">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="03314-110">L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="03314-110">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="03314-111">`Range.copyFrom`dispose de trois paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="03314-111">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="03314-112">`copyType` spécifie les données copiées de la source vers la destination.</span><span class="sxs-lookup"><span data-stu-id="03314-112">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="03314-113">`Excel.RangeCopyType.formulas` transfère les formules dans les cellules sources et conserve le positionnement relatif des plages de ces formules.</span><span class="sxs-lookup"><span data-stu-id="03314-113">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="03314-114">Les entrées autres que des formules sont copiées telles quelles.</span><span class="sxs-lookup"><span data-stu-id="03314-114">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="03314-115">`Excel.RangeCopyType.values` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="03314-115">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="03314-116">`Excel.RangeCopyType.formats` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.</span><span class="sxs-lookup"><span data-stu-id="03314-116">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="03314-117">`Excel.RangeCopyType.all` (option par défaut) copie les données et la mise en forme, en conservant les formules des cellules si elles sont trouvées.</span><span class="sxs-lookup"><span data-stu-id="03314-117">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="03314-118">`skipBlanks` définit si les cellules vides sont copiées dans la destination.</span><span class="sxs-lookup"><span data-stu-id="03314-118">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="03314-119">Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.</span><span class="sxs-lookup"><span data-stu-id="03314-119">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="03314-120">Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination.</span><span class="sxs-lookup"><span data-stu-id="03314-120">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="03314-121">La valeur par défaut est false.</span><span class="sxs-lookup"><span data-stu-id="03314-121">The default is false.</span></span>

<span data-ttu-id="03314-122">`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.</span><span class="sxs-lookup"><span data-stu-id="03314-122">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="03314-123">Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.</span><span class="sxs-lookup"><span data-stu-id="03314-123">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="03314-124">L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="03314-124">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-copied-and-pasted"></a><span data-ttu-id="03314-125">Données avant que la plage ne soit copiée et copiée</span><span class="sxs-lookup"><span data-stu-id="03314-125">Data before range is copied and pasted</span></span>

![Données dans Excel la méthode de copie de plage a été exécuté.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a><span data-ttu-id="03314-127">Données une fois la plage copiée et copiée</span><span class="sxs-lookup"><span data-stu-id="03314-127">Data after range is copied and pasted</span></span>

![Données dans Excel une fois que la méthode de copie de plage a été exécuté.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a><span data-ttu-id="03314-129">Couper et coller (déplacer) des cellules</span><span class="sxs-lookup"><span data-stu-id="03314-129">Cut and paste (move) cells</span></span>

<span data-ttu-id="03314-130">La [méthode Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) déplace les cellules vers un nouvel emplacement dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="03314-130">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="03314-131">Ce comportement de déplacement de cellule fonctionne [](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) de la même manière  que lorsque les cellules sont déplacées en faisant glisser la bordure de la plage ou lors de l’action Couper **et** coller.</span><span class="sxs-lookup"><span data-stu-id="03314-131">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="03314-132">La mise en forme et les valeurs de la plage sont déplacées vers l’emplacement spécifié en tant que `destinationRange` paramètre.</span><span class="sxs-lookup"><span data-stu-id="03314-132">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="03314-133">L’exemple de code suivant déplace une plage avec la `Range.moveTo` méthode.</span><span class="sxs-lookup"><span data-stu-id="03314-133">The following code sample moves a range with the `Range.moveTo` method.</span></span> <span data-ttu-id="03314-134">Notez que si la plage de destination est plus petite que la source, elle sera étendue pour englober le contenu source.</span><span class="sxs-lookup"><span data-stu-id="03314-134">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="03314-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="03314-135">See also</span></span>

- [<span data-ttu-id="03314-136">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="03314-136">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="03314-137">Utiliser des cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="03314-137">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="03314-138">Supprimer les doublons à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="03314-138">Remove duplicates using the Excel JavaScript API</span></span>](excel-add-ins-ranges-remove-duplicates.md)
- [<span data-ttu-id="03314-139">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="03314-139">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
