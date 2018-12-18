---
title: Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 42b1127580c46120d337553fdb86a19a78b37567
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283792"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="1a456-102">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="1a456-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="1a456-103">Cet article génère des informations dans[ Utiliser des plages à l’aide de l’API JavaScript Excel (fondamental)](excel-add-ins-ranges.md) en fournissant les exemples de code qui affichent la manière d’exécuter plus de tâches avancées avec des plages à l’aide de l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="1a456-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="1a456-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Range**, reportez-vous à la rubrique [Objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="1a456-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="1a456-105">Utiliser des dates à l’aide de plug-in Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="1a456-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="1a456-106">La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs.</span><span class="sxs-lookup"><span data-stu-id="1a456-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="1a456-107">Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel.</span><span class="sxs-lookup"><span data-stu-id="1a456-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="1a456-108">Il s’agit du même format que la[fonction NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)renvoie.</span><span class="sxs-lookup"><span data-stu-id="1a456-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="1a456-109">Le code suivant affiche la manière d’établir la plage à**B4**vers un horodateur du moment:</span><span class="sxs-lookup"><span data-stu-id="1a456-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1a456-110">Il s’agit d’une technique similaire pour récupérer la date de la cellule et la convertir en un moment ou autre format, comme démontré dans le code suivant:</span><span class="sxs-lookup"><span data-stu-id="1a456-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1a456-111">Votre complément devra mettre en forme les plages pour afficher les dates dans une forme plus lisible.</span><span class="sxs-lookup"><span data-stu-id="1a456-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="1a456-112">L’exemple de`"[$-409]m/d/yy h:mm AM/PM;@"`affiche une heure comme «12/3/18 3:57 PM».</span><span class="sxs-lookup"><span data-stu-id="1a456-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="1a456-113">Pour plus d’informations concernant les formats de date et d’heure , veuillez consulter les «Instructions relatifs aux formats de date et heure» dans l’article[ Instructions revoir afin de personnaliser le format numérique](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="1a456-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="copy-and-paste"></a><span data-ttu-id="1a456-114">Copier et coller</span><span class="sxs-lookup"><span data-stu-id="1a456-114">Copy and Paste</span></span>

> [!NOTE]
> <span data-ttu-id="1a456-115">La fonction`Range.copyFrom` est actuellement disponible uniquement en préversion publique (bêta).</span><span class="sxs-lookup"><span data-stu-id="1a456-115">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="1a456-116">Pour utiliser cette fonctionnalité, vous devez utiliser la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="1a456-116">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="1a456-117">Si vous utilisez TypeScript ou si votre éditeur de code utilise des fichiers de définition de type TypeScript pour IntelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="1a456-117">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="1a456-118">La fonction`copyFrom`de la plage reproduit le comportement de copier-coller de l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="1a456-118">Range’s copyFrom function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="1a456-119">L’objet plage sur lequel`copyFrom`est appelé est la destination.</span><span class="sxs-lookup"><span data-stu-id="1a456-119">The range object that copyFrom is called on is the destination.</span></span>
<span data-ttu-id="1a456-120">La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.</span><span class="sxs-lookup"><span data-stu-id="1a456-120">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="1a456-121">L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="1a456-121">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1a456-122">`Range.copyFrom`dispose de trois paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="1a456-122">Range.copyFrom has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="1a456-123">`copyType` spécifie les données copiées de la source vers la destination.</span><span class="sxs-lookup"><span data-stu-id="1a456-123">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="1a456-124">`"Formulas"` transfère les formules dans les cellules sources en préservant le positionnement relatif des plages de ces formules.</span><span class="sxs-lookup"><span data-stu-id="1a456-124">`"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="1a456-125">Les entrées autres que des formules sont copiées telles quelles.</span><span class="sxs-lookup"><span data-stu-id="1a456-125">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="1a456-126">`"Values"` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="1a456-126">`"Values"` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="1a456-127">`"Formats"` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.</span><span class="sxs-lookup"><span data-stu-id="1a456-127">`"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="1a456-128">`"All"` (option par défaut) copie les données et la mise en forme, en conservant les formules éventuelles des cellules.</span><span class="sxs-lookup"><span data-stu-id="1a456-128">`"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="1a456-129">`skipBlanks` définit si les cellules vides sont copiées dans la destination.</span><span class="sxs-lookup"><span data-stu-id="1a456-129">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="1a456-130">Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.</span><span class="sxs-lookup"><span data-stu-id="1a456-130">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="1a456-131">Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination.</span><span class="sxs-lookup"><span data-stu-id="1a456-131">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="1a456-132">La valeur par défaut est false.</span><span class="sxs-lookup"><span data-stu-id="1a456-132">The default is false.</span></span>

<span data-ttu-id="1a456-133">`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.</span><span class="sxs-lookup"><span data-stu-id="1a456-133">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="1a456-134">Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.</span><span class="sxs-lookup"><span data-stu-id="1a456-134">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="1a456-135">L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="1a456-135">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="1a456-136">*Avant l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="1a456-136">*Before the preceeding function has been run.*</span></span>

![Données dans Excel avant exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="1a456-138">*Après l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="1a456-138">*After the preceeding function has been run.*</span></span>

![Données dans Excel après exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="1a456-140">Supprimer les doublons</span><span class="sxs-lookup"><span data-stu-id="1a456-140">Remove duplicates</span></span>

> [!NOTE]
> <span data-ttu-id="1a456-141">La fonction`removeDuplicates` de l’objet de la plage est actuellement disponible uniquement en préversion publique (bêta).</span><span class="sxs-lookup"><span data-stu-id="1a456-141">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="1a456-142">Pour utiliser cette fonctionnalité, vous devez utiliser la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="1a456-142">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="1a456-143">Si vous utilisez TypeScript ou si votre éditeur de code utilise des fichiers de définition de type TypeScript pour IntelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="1a456-143">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="1a456-144">La fonction`removeDuplicates`de l’objet de la plage retire les rangées avec les entrées en doublon dans les colonnes spécifiées.</span><span class="sxs-lookup"><span data-stu-id="1a456-144">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="1a456-145">La fonction circule à travers chaque rangée de la plage de l’index à la valeur la plus basse à l’index à la valeur la plus haute de la plage ( du haut vers le bas).</span><span class="sxs-lookup"><span data-stu-id="1a456-145">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="1a456-146">Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage.</span><span class="sxs-lookup"><span data-stu-id="1a456-146">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="1a456-147">Les rangées de la plage en-dessous de la rangée supprimée sont déplacées.</span><span class="sxs-lookup"><span data-stu-id="1a456-147">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="1a456-148">`removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.</span><span class="sxs-lookup"><span data-stu-id="1a456-148">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="1a456-149">`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons.</span><span class="sxs-lookup"><span data-stu-id="1a456-149">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="1a456-150">Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="1a456-150">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="1a456-151">La fonction prend également un paramètre booléen qui spécifie si la première rangée est un-tête.</span><span class="sxs-lookup"><span data-stu-id="1a456-151">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="1a456-152">Lorsque**true**, la rangée du dessus est ignorée lorsque les doublons sont pris en considération.</span><span class="sxs-lookup"><span data-stu-id="1a456-152">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="1a456-153">La fonction`removeDuplicates`renvoie un objet`RemoveDuplicatesResult` qui spécifie le nombre de rangée retirées et le nombre de rangées uniques restantes.</span><span class="sxs-lookup"><span data-stu-id="1a456-153">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="1a456-154">Lors de l’usage d’une fonction`removeDuplicates`de la plage, gardez ce qui suit à l’esprit:</span><span class="sxs-lookup"><span data-stu-id="1a456-154">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="1a456-155">`removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction.</span><span class="sxs-lookup"><span data-stu-id="1a456-155">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="1a456-156">Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.</span><span class="sxs-lookup"><span data-stu-id="1a456-156">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="1a456-157">Les cellules vides ne sont pas ignorées par`removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="1a456-157">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="1a456-158">La valeur d’une cellule vide est traitée comme toute autre valeur.</span><span class="sxs-lookup"><span data-stu-id="1a456-158">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="1a456-159">Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="1a456-159">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="1a456-160">L’exemple suivant affiche la suppression des entrées avec des valeurs de doublons dans la première colonne.</span><span class="sxs-lookup"><span data-stu-id="1a456-160">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
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

<span data-ttu-id="1a456-161">*Avant l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="1a456-161">*Before the preceeding function has been run.*</span></span>

![Données dans Excel avant exécution de la méthode de copie de suppression de la plage.](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="1a456-163">*Après l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="1a456-163">*After the preceeding function has been run.*</span></span>

![Données dans Excel après exécution de la méthode de copie de suppression de la plage.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="1a456-165">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1a456-165">See also</span></span>

- [<span data-ttu-id="1a456-166">Utiliser les plages à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="1a456-166">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="1a456-167">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1a456-167">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)