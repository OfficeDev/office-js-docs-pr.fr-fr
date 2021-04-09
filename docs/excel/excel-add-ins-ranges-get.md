---
title: Obtenir une plage à l’aide de l’API JavaScript pour Excel
description: Découvrez comment récupérer une plage à l’aide de l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652848"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="0fd13-103">Obtenir une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="0fd13-103">Get a range using the Excel JavaScript API</span></span>

<span data-ttu-id="0fd13-104">Cet article fournit des exemples qui montrent différentes façons d’obtenir une plage dans une feuille de calcul à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="0fd13-104">This article provides examples that show different ways to get a range within a worksheet using the Excel JavaScript API.</span></span> <span data-ttu-id="0fd13-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="0fd13-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a><span data-ttu-id="0fd13-106">Obtenir une plage en fonction d’une adresse</span><span class="sxs-lookup"><span data-stu-id="0fd13-106">Get range by address</span></span>

<span data-ttu-id="0fd13-107">L’exemple de code suivant obtient la plage avec l’adresse **B2:C5** à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit un `address` message dans la console.</span><span class="sxs-lookup"><span data-stu-id="0fd13-107">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a><span data-ttu-id="0fd13-108">Obtenir une plage en fonction d’un nom</span><span class="sxs-lookup"><span data-stu-id="0fd13-108">Get range by name</span></span>

<span data-ttu-id="0fd13-109">L’exemple de code suivant obtient la plage nommée à partir de la feuille de calcul nommée Sample, charge sa propriété et écrit `MyRange` un message dans la  `address` console.</span><span class="sxs-lookup"><span data-stu-id="0fd13-109">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a><span data-ttu-id="0fd13-110">Obtenir une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="0fd13-110">Get used range</span></span>

<span data-ttu-id="0fd13-111">L’exemple de code suivant obtient la plage utilisée à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit `address` un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="0fd13-111">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="0fd13-112">La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté.</span><span class="sxs-lookup"><span data-stu-id="0fd13-112">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="0fd13-113">Si la feuille de calcul entière est vide, la méthode renvoie une plage qui se compose uniquement de `getUsedRange()` la cellule supérieure gauche.</span><span class="sxs-lookup"><span data-stu-id="0fd13-113">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a><span data-ttu-id="0fd13-114">Obtenir l’intégralité d’une plage</span><span class="sxs-lookup"><span data-stu-id="0fd13-114">Get entire range</span></span>

<span data-ttu-id="0fd13-115">L’exemple de code suivant obtient l’ensemble de la plage de feuille de calcul à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit un `address` message dans la console.</span><span class="sxs-lookup"><span data-stu-id="0fd13-115">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="0fd13-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0fd13-116">See also</span></span>

- [<span data-ttu-id="0fd13-117">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="0fd13-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0fd13-118">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="0fd13-118">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="0fd13-119">Insérer une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="0fd13-119">Insert a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-insert.md)
