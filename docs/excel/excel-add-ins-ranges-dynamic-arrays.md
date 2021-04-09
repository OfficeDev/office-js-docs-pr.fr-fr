---
title: Gérer les tableaux dynamiques et la plage qui se débordent à l’aide de l’API JavaScript pour Excel
description: Découvrez comment gérer les tableaux dynamiques et la plage qui se débordent avec l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652856"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a><span data-ttu-id="e6579-103">Gérer les tableaux dynamiques et les débordements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e6579-103">Handle dynamic arrays and spilling using the Excel JavaScript API</span></span>

<span data-ttu-id="e6579-104">Cet article fournit un exemple de code qui gère les tableaux dynamiques et les étendues à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="e6579-104">This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API.</span></span> <span data-ttu-id="e6579-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="e6579-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="dynamic-arrays"></a><span data-ttu-id="e6579-106">Tableaux dynamiques</span><span class="sxs-lookup"><span data-stu-id="e6579-106">Dynamic arrays</span></span>

<span data-ttu-id="e6579-107">Certaines formules Excel retournent [des tableaux dynamiques.](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)</span><span class="sxs-lookup"><span data-stu-id="e6579-107">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="e6579-108">Ceux-ci remplissent les valeurs de plusieurs cellules en dehors de la cellule d’origine de la formule.</span><span class="sxs-lookup"><span data-stu-id="e6579-108">These fill the values of multiple cells outside of the formula's original cell.</span></span> <span data-ttu-id="e6579-109">Cette valeur de dépassement est appelée « débordement ».</span><span class="sxs-lookup"><span data-stu-id="e6579-109">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="e6579-110">Votre add-in peut trouver la plage utilisée pour un débordement avec la [méthode Range.getSpillingToRange.](/javascript/api/excel/excel.range#getspillingtorange--)</span><span class="sxs-lookup"><span data-stu-id="e6579-110">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="e6579-111">Il existe également [une version \*OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="e6579-111">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="e6579-112">L’exemple suivant montre une formule de base qui copie le contenu d’une plage dans une cellule, qui se renverse dans les cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="e6579-112">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="e6579-113">Le add-in enregistre ensuite la plage qui contient le débordement.</span><span class="sxs-lookup"><span data-stu-id="e6579-113">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a><span data-ttu-id="e6579-114">Étendue de plage</span><span class="sxs-lookup"><span data-stu-id="e6579-114">Range spilling</span></span>

<span data-ttu-id="e6579-115">Recherchez la cellule responsable du débordement dans une cellule donnée à l’aide de la [méthode Range.getSpillParent.](/javascript/api/excel/excel.range#getspillparent--)</span><span class="sxs-lookup"><span data-stu-id="e6579-115">Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="e6579-116">Notez que `getSpillParent` fonctionne uniquement lorsque l’objet de plage est une seule cellule.</span><span class="sxs-lookup"><span data-stu-id="e6579-116">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="e6579-117">L’appel sur une plage avec plusieurs cellules entraîne une erreur en cours de thrown (ou une plage `getSpillParent` null renvoyée pour `Range.getSpillParentOrNullObject` ).</span><span class="sxs-lookup"><span data-stu-id="e6579-117">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="e6579-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e6579-118">See also</span></span>

- [<span data-ttu-id="e6579-119">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="e6579-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="e6579-120">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e6579-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="e6579-121">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="e6579-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
