---
title: Lire ou écrire dans une plage non limite à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript pour Excel pour lire ou écrire dans une plage non limite.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652820"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a><span data-ttu-id="00f92-103">Lire ou écrire dans une plage non limite à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="00f92-103">Read or write to an unbounded range using the Excel JavaScript API</span></span>

<span data-ttu-id="00f92-104">Cet article explique comment lire et écrire dans une plage non limite avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="00f92-104">This article describes how to read and write to an unbounded range with the Excel JavaScript API.</span></span> <span data-ttu-id="00f92-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="00f92-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

<span data-ttu-id="00f92-106">Une adresse de plage non limite est une adresse de plage qui spécifie des colonnes entières ou des lignes entières.</span><span class="sxs-lookup"><span data-stu-id="00f92-106">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="00f92-107">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="00f92-107">For example:</span></span>

- <span data-ttu-id="00f92-108">Adresses de plage composées de colonnes entières :</span><span class="sxs-lookup"><span data-stu-id="00f92-108">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="00f92-109">Adresses de plage composées de lignes entières :</span><span class="sxs-lookup"><span data-stu-id="00f92-109">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a><span data-ttu-id="00f92-110">Lire une plage non liée</span><span class="sxs-lookup"><span data-stu-id="00f92-110">Read an unbounded range</span></span>

<span data-ttu-id="00f92-p103">Lorsque l’API effectue une demande de récupération d’une plage non liée (par exemple, `getRange('C:C')`), la réponse contient des valeurs `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`. Les autres propriétés de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non liée.</span><span class="sxs-lookup"><span data-stu-id="00f92-p103">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

## <a name="write-to-an-unbounded-range"></a><span data-ttu-id="00f92-113">Écrire dans une plage non liée</span><span class="sxs-lookup"><span data-stu-id="00f92-113">Write to an unbounded range</span></span>

<span data-ttu-id="00f92-114">Vous ne pouvez pas définir de propriétés au niveau de la cellule telles que , et sur une plage non limite, car la demande d’entrée `values` `numberFormat` est trop `formula` grande.</span><span class="sxs-lookup"><span data-stu-id="00f92-114">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="00f92-115">Par exemple, l’exemple de code suivant n’est pas valide, car il tente de spécifier une plage `values` non limite.</span><span class="sxs-lookup"><span data-stu-id="00f92-115">For example, the following code example is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="00f92-116">L’API renvoie une erreur si vous tentez de définir des propriétés au niveau de la cellule pour une plage non limite.</span><span class="sxs-lookup"><span data-stu-id="00f92-116">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a><span data-ttu-id="00f92-117">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="00f92-117">See also</span></span>

- [<span data-ttu-id="00f92-118">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="00f92-118">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="00f92-119">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="00f92-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="00f92-120">Lire ou écrire dans une grande plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="00f92-120">Read or write to a large range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-large.md)
- [<span data-ttu-id="00f92-121">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="00f92-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
