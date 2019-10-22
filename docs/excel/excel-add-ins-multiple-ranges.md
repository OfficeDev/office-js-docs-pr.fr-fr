---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: ''
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: a327b6c379884107f5e00c0663ecfa6c71b8097f
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33620044"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a><span data-ttu-id="c48fd-102">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="c48fd-102">Work with multiple ranges simultaneously in Excel add-ins</span></span>

<span data-ttu-id="c48fd-103">La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages.</span><span class="sxs-lookup"><span data-stu-id="c48fd-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="c48fd-104">Les plages n’ont pas besoin d’être contigus.</span><span class="sxs-lookup"><span data-stu-id="c48fd-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="c48fd-105">En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.</span><span class="sxs-lookup"><span data-stu-id="c48fd-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

## <a name="rangeareas"></a><span data-ttu-id="c48fd-106">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c48fd-106">RangeAreas</span></span>

<span data-ttu-id="c48fd-107">Un ensemble de plages (éventuellement discontinues) est représenté par un objet [RangeAreas](/javascript/api/excel/excel.rangeareas) .</span><span class="sxs-lookup"><span data-stu-id="c48fd-107">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="c48fd-108">Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:</span><span class="sxs-lookup"><span data-stu-id="c48fd-108">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="c48fd-109">Les types de données pour les propriétés et le comportement des méthodes et des getters.</span><span class="sxs-lookup"><span data-stu-id="c48fd-109">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="c48fd-110">Les types de données de paramètres et des comportements de la méthode.</span><span class="sxs-lookup"><span data-stu-id="c48fd-110">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="c48fd-111">Les types de données de méthodes renvoient des valeurs.</span><span class="sxs-lookup"><span data-stu-id="c48fd-111">The data types of method return values.</span></span>

<span data-ttu-id="c48fd-112">Quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="c48fd-112">Some examples:</span></span>

- <span data-ttu-id="c48fd-113">`RangeAreas` a une`address` propriété qui renvoie une chaîne séparée par des virgules de plage d’adresses, au lieu d’une adresse comme avec la `Range.address` propriété.</span><span class="sxs-lookup"><span data-stu-id="c48fd-113">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="c48fd-114">`RangeAreas` a une`dataValidation` propriété qui renvoie un`DataValidation` objet qui représente la validation des données de toutes les plages dans la `RangeAreas`, s’il est cohérent.</span><span class="sxs-lookup"><span data-stu-id="c48fd-114">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="c48fd-115">La propriété est`null` si les objets`DataValidation` identiques ne sont pas appliqués à toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-115">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c48fd-116">Il s’agit d’un principe général, mais pas universel avec l’`RangeAreas` objet: *si une propriété ne comporte pas de valeurs cohérentes sur tous les plages dans la`RangeAreas`, cela signifie`null`.*</span><span class="sxs-lookup"><span data-stu-id="c48fd-116">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="c48fd-117">Voir[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) pour plus d’informations et quelques exceptions.</span><span class="sxs-lookup"><span data-stu-id="c48fd-117">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="c48fd-118">`RangeAreas.cellCount` Obtient le nombre total de cellules dans toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-118">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c48fd-119">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-119">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c48fd-120">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` retourner un autre`RangeAreas` objet qui représente toutes les colonnes (ou lignes) dans toutes les plages dans la `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-120">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c48fd-121">Par exemple, si le`RangeAreas` représente « A1 : C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un`RangeAreas` objet qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="c48fd-121">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="c48fd-122">`RangeAreas.copyFrom` peut prendre soit un`Range` ou d’un`RangeAreas` paramètre représentant la ou les plage(s) source de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="c48fd-122">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="c48fd-123">La liste complète des membres plage sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c48fd-123">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="c48fd-124">Propriétés</span><span class="sxs-lookup"><span data-stu-id="c48fd-124">Properties</span></span>

<span data-ttu-id="c48fd-125">Être familiarisé avec[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) avant d’écrire de code qui lit les propriétés répertoriées.</span><span class="sxs-lookup"><span data-stu-id="c48fd-125">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="c48fd-126">Il existe des subtilités sur ce qui est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="c48fd-126">There are subtleties to what gets returned.</span></span>

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a><span data-ttu-id="c48fd-127">Méthodes</span><span class="sxs-lookup"><span data-stu-id="c48fd-127">Methods</span></span>

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- <span data-ttu-id="c48fd-128">`getOffsetRange()`(nommé `getOffsetRangeAreas` sur l' `RangeAreas` objet)</span><span class="sxs-lookup"><span data-stu-id="c48fd-128">`getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)</span></span>
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- <span data-ttu-id="c48fd-129">`getUsedRange()`(nommé `getUsedRangeAreas` sur l' `RangeAreas` objet)</span><span class="sxs-lookup"><span data-stu-id="c48fd-129">`getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)</span></span>
- <span data-ttu-id="c48fd-130">`getUsedRangeOrNullObject()`(nommé `getUsedRangeAreasOrNullObject` sur l' `RangeAreas` objet)</span><span class="sxs-lookup"><span data-stu-id="c48fd-130">`getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)</span></span>
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="c48fd-131">Méthodes et propriétés propres à une langue RangeArea</span><span class="sxs-lookup"><span data-stu-id="c48fd-131">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="c48fd-132">Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet.</span><span class="sxs-lookup"><span data-stu-id="c48fd-132">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="c48fd-133">Ce qui est une sélection de certains d’entre eux :</span><span class="sxs-lookup"><span data-stu-id="c48fd-133">The following is a selection of them:</span></span>

- <span data-ttu-id="c48fd-134">`areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet.</span><span class="sxs-lookup"><span data-stu-id="c48fd-134">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="c48fd-135">L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel.</span><span class="sxs-lookup"><span data-stu-id="c48fd-135">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="c48fd-136">Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="c48fd-136">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="c48fd-137">`areaCount`: Le nombre total de plages dans le`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-137">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c48fd-138">`getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-138">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="c48fd-139">Créer l’objet RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c48fd-139">Create RangeAreas</span></span>

<span data-ttu-id="c48fd-140">Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:</span><span class="sxs-lookup"><span data-stu-id="c48fd-140">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="c48fd-141">Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="c48fd-141">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="c48fd-142">Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="c48fd-142">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="c48fd-143">Appel `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-143">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="c48fd-144">Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.</span><span class="sxs-lookup"><span data-stu-id="c48fd-144">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="c48fd-145">Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-145">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="c48fd-146">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-146">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="c48fd-147">Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-147">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="c48fd-148">N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-148">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="c48fd-149">Cela mènera à un comportement indésirable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="c48fd-149">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="c48fd-150">Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas.</span><span class="sxs-lookup"><span data-stu-id="c48fd-150">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="c48fd-151">Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-151">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="c48fd-152">De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet*soit*supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore.</span><span class="sxs-lookup"><span data-stu-id="c48fd-152">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="c48fd-153">Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.</span><span class="sxs-lookup"><span data-stu-id="c48fd-153">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="c48fd-154">Définir les propriétés sur plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="c48fd-154">Set properties on multiple ranges</span></span>

<span data-ttu-id="c48fd-155">Paramétrer une propriété sur un objet `RangeAreas` établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-155">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="c48fd-156">Ce qui suit est un exemple de paramétrage d’une propriété sur des plages multiples.</span><span class="sxs-lookup"><span data-stu-id="c48fd-156">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="c48fd-157">La fonction surligne les plages**F3:F5** and **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="c48fd-157">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="c48fd-158">Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution.</span><span class="sxs-lookup"><span data-stu-id="c48fd-158">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="c48fd-159">Certains des scénarios dans lesquels ceci peut s’appliquer incluent:</span><span class="sxs-lookup"><span data-stu-id="c48fd-159">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="c48fd-160">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="c48fd-160">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="c48fd-161">Le code s’exécute dans le contexte de données importées où le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="c48fd-161">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="c48fd-162">Obtenir des cellules spéciales à partir de plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="c48fd-162">Get special cells from multiple ranges</span></span>

<span data-ttu-id="c48fd-163">Les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` sur l’objet `RangeAreas` fonctionnent de manière analogue aux méthodes du même nom sur l’objet `Range`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-163">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="c48fd-164">Ces méthodes retournent les cellules disposant de la caractéristique spécifiée à partir de toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-164">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="c48fd-165">Voir la section [Trouver des cellules spéciales dans une plage](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) pour plus d’informations sur les cellules spéciales.</span><span class="sxs-lookup"><span data-stu-id="c48fd-165">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) section for more details on special cells.</span></span>

<span data-ttu-id="c48fd-166">Lors de l’appel de la méthode `getSpecialCells` ou `getSpecialCellsOrNullObject` sur un objet `RangeAreas` :</span><span class="sxs-lookup"><span data-stu-id="c48fd-166">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="c48fd-167">si vous passez `Excel.SpecialCellType.sameConditionalFormat` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même mise en forme conditionnelle que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-167">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="c48fd-168">Si vous passez `Excel.SpecialCellType.sameDataValidation` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même règle de validation des données que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-168">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="c48fd-169">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c48fd-169">Read properties of RangeAreas</span></span>

<span data-ttu-id="c48fd-170">La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-170">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="c48fd-171">La règle générales est que si une valeur consistante*peut*être renvoyée, elle sera renvoyée.</span><span class="sxs-lookup"><span data-stu-id="c48fd-171">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="c48fd-172">Par exemple, dans le code suivant, le code RGB pour rose (`#FFC0CB`) et`true`sera connecté à la console car les deux plages dans l’objet`RangeAreas` dispose d’un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="c48fd-172">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

<span data-ttu-id="c48fd-173">Les choses se compliquent lorsque la consistance est impossible.</span><span class="sxs-lookup"><span data-stu-id="c48fd-173">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="c48fd-174">Le comportement de propriétés`RangeAreas` suit ces trois principes:</span><span class="sxs-lookup"><span data-stu-id="c48fd-174">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="c48fd-175">Une propriété booléenne d’un objet`RangeAreas` renvoie`false`à moins que la propriété soit vraie pour toutes les plages membres.</span><span class="sxs-lookup"><span data-stu-id="c48fd-175">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="c48fd-176">Les propriétés non-booléennes, avec l’exception de la propriété`address`renvoie`null`à moins que la propriété correspondante sur toutes les plages membres dispose de la même valeur.</span><span class="sxs-lookup"><span data-stu-id="c48fd-176">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="c48fd-177">La propriété`address`renvoie une chaîne délimitée à virgule des adresses des plages membres.</span><span class="sxs-lookup"><span data-stu-id="c48fd-177">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="c48fd-178">Par exemple, le code suivante crée un`RangeAreas`dans lequel seule une plage est une colonne entière et seule une est remplie de rose.</span><span class="sxs-lookup"><span data-stu-id="c48fd-178">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="c48fd-179">La console s’affichera`null`pour un remplissage de couleur,`false`pour la propriété`isEntireRow` et «Sheet1!F3:F5, Sheet1!H:H» (en présumant que la feuille de calcule soit «Sheet1») pour la propriété`address`.</span><span class="sxs-lookup"><span data-stu-id="c48fd-179">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a><span data-ttu-id="c48fd-180">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c48fd-180">See also</span></span>

- [<span data-ttu-id="c48fd-181">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c48fd-181">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="c48fd-182">Utilisation de plages à l’aide de l’API JavaScript pour Excel (basique)</span><span class="sxs-lookup"><span data-stu-id="c48fd-182">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="c48fd-183">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="c48fd-183">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
