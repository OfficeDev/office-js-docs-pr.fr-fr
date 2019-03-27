---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: ''
ms.date: 02/20/2019
localization_priority: Normal
ms.openlocfilehash: d284f3cb12153b1c05cff0b7f7a66bbf98961443
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870820"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="e91c6-102">Travailler simultanément avec plusieurs plages dans des compléments Excel (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="e91c6-102">Work with multiple ranges simultaneously in Excel add-ins (preview)</span></span>

<span data-ttu-id="e91c6-103">La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages.</span><span class="sxs-lookup"><span data-stu-id="e91c6-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="e91c6-104">Les plages n’ont pas besoin d’être contigus.</span><span class="sxs-lookup"><span data-stu-id="e91c6-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="e91c6-105">En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.</span><span class="sxs-lookup"><span data-stu-id="e91c6-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="e91c6-106">Les APIs décrits dans cet article nécessitent**la version Office 2016 «Démarrer en un Clic» 1809 Build 10820.20000**ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="e91c6-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="e91c6-107">(Vous devrez peut-être rejoindre le [programme Office](https://products.office.com/office-insider) Insider pour obtenir une build appropriée.)[!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]</span><span class="sxs-lookup"><span data-stu-id="e91c6-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.)  [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]</span></span>

## <a name="rangeareas"></a><span data-ttu-id="e91c6-108">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e91c6-108">RangeAreas</span></span>

<span data-ttu-id="e91c6-109">Un ensemble de plages (éventuellement discontinues) est représenté par un objet [RangeAreas](/javascript/api/excel/excel.rangeareas) .</span><span class="sxs-lookup"><span data-stu-id="e91c6-109">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="e91c6-110">Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:</span><span class="sxs-lookup"><span data-stu-id="e91c6-110">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="e91c6-111">Les types de données pour les propriétés et le comportement des méthodes et des getters.</span><span class="sxs-lookup"><span data-stu-id="e91c6-111">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="e91c6-112">Les types de données de paramètres et des comportements de la méthode.</span><span class="sxs-lookup"><span data-stu-id="e91c6-112">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="e91c6-113">Les types de données de méthodes renvoient des valeurs.</span><span class="sxs-lookup"><span data-stu-id="e91c6-113">The data types of method return values.</span></span>

<span data-ttu-id="e91c6-114">Quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="e91c6-114">Some examples:</span></span>

- <span data-ttu-id="e91c6-115">`RangeAreas` a une`address` propriété qui renvoie une chaîne séparée par des virgules de plage d’adresses, au lieu d’une adresse comme avec la `Range.address` propriété.</span><span class="sxs-lookup"><span data-stu-id="e91c6-115">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="e91c6-116">`RangeAreas` a une`dataValidation` propriété qui renvoie un`DataValidation` objet qui représente la validation des données de toutes les plages dans la `RangeAreas`, s’il est cohérent.</span><span class="sxs-lookup"><span data-stu-id="e91c6-116">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="e91c6-117">La propriété est`null` si les objets`DataValidation` identiques ne sont pas appliqués à toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-117">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="e91c6-118">Il s’agit d’un principe général, mais pas universel avec l’`RangeAreas` objet: *si une propriété ne comporte pas de valeurs cohérentes sur tous les plages dans la`RangeAreas`, cela signifie`null`.*</span><span class="sxs-lookup"><span data-stu-id="e91c6-118">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="e91c6-119">Voir[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) pour plus d’informations et quelques exceptions.</span><span class="sxs-lookup"><span data-stu-id="e91c6-119">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="e91c6-120">`RangeAreas.cellCount` Obtient le nombre total de cellules dans toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-120">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e91c6-121">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-121">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e91c6-122">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` retourner un autre`RangeAreas` objet qui représente toutes les colonnes (ou lignes) dans toutes les plages dans la `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-122">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="e91c6-123">Par exemple, si le`RangeAreas` représente « A1 : C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un`RangeAreas` objet qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="e91c6-123">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="e91c6-124">`RangeAreas.copyFrom` peut prendre soit un`Range` ou d’un`RangeAreas` paramètre représentant la ou les plage(s) source de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="e91c6-124">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="e91c6-125">La liste complète des membres plage sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e91c6-125">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="e91c6-126">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e91c6-126">Properties</span></span>

<span data-ttu-id="e91c6-127">Être familiarisé avec[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) avant d’écrire de code qui lit les propriétés répertoriées.</span><span class="sxs-lookup"><span data-stu-id="e91c6-127">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="e91c6-128">Il existe des subtilités sur ce qui est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="e91c6-128">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="e91c6-129">address</span><span class="sxs-lookup"><span data-stu-id="e91c6-129">address</span></span>
- <span data-ttu-id="e91c6-130">addressLocal</span><span class="sxs-lookup"><span data-stu-id="e91c6-130">addressLocal</span></span>
- <span data-ttu-id="e91c6-131">cellCount</span><span class="sxs-lookup"><span data-stu-id="e91c6-131">cellCount</span></span>
- <span data-ttu-id="e91c6-132">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="e91c6-132">conditionalFormats</span></span>
- <span data-ttu-id="e91c6-133">context</span><span class="sxs-lookup"><span data-stu-id="e91c6-133">context</span></span>
- <span data-ttu-id="e91c6-134">dataValidation</span><span class="sxs-lookup"><span data-stu-id="e91c6-134">dataValidation</span></span>
- <span data-ttu-id="e91c6-135">format</span><span class="sxs-lookup"><span data-stu-id="e91c6-135">format</span></span>
- <span data-ttu-id="e91c6-136">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="e91c6-136">isEntireColumn</span></span>
- <span data-ttu-id="e91c6-137">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="e91c6-137">isEntireRow</span></span>
- <span data-ttu-id="e91c6-138">style</span><span class="sxs-lookup"><span data-stu-id="e91c6-138">style</span></span>
- <span data-ttu-id="e91c6-139">worksheet</span><span class="sxs-lookup"><span data-stu-id="e91c6-139">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="e91c6-140">Méthodes</span><span class="sxs-lookup"><span data-stu-id="e91c6-140">Methods</span></span>

<span data-ttu-id="e91c6-141">Les méthodes de plage dans l’aperçu sont marquées.</span><span class="sxs-lookup"><span data-stu-id="e91c6-141">Range methods in preview are marked.</span></span>

- <span data-ttu-id="e91c6-142">calculate()</span><span class="sxs-lookup"><span data-stu-id="e91c6-142">calculate()</span></span>
- <span data-ttu-id="e91c6-143">clear()</span><span class="sxs-lookup"><span data-stu-id="e91c6-143">clear()</span></span>
- <span data-ttu-id="e91c6-144">convertDataTypeToText() (preview)</span><span class="sxs-lookup"><span data-stu-id="e91c6-144">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="e91c6-145">convertToLinkedDataType() (preview)</span><span class="sxs-lookup"><span data-stu-id="e91c6-145">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="e91c6-146">copyFrom() (preview)</span><span class="sxs-lookup"><span data-stu-id="e91c6-146">copyFrom() (preview)</span></span>
- <span data-ttu-id="e91c6-147">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="e91c6-147">getEntireColumn()</span></span>
- <span data-ttu-id="e91c6-148">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="e91c6-148">getEntireRow()</span></span>
- <span data-ttu-id="e91c6-149">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="e91c6-149">getIntersection()</span></span>
- <span data-ttu-id="e91c6-150">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="e91c6-150">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="e91c6-151">getOffsetRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e91c6-151">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="e91c6-152">getSpecialCells() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="e91c6-152">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="e91c6-153">getSpecialCellsOrNullObject() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="e91c6-153">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="e91c6-154">getTables() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="e91c6-154">getTables() (preview)</span></span>
- <span data-ttu-id="e91c6-155">getUsedRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e91c6-155">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="e91c6-156">getUsedRangeOrNullObject() (appelé getUsedRangeAreasOrNullObject sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e91c6-156">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="e91c6-157">load()</span><span class="sxs-lookup"><span data-stu-id="e91c6-157">load()</span></span>
- <span data-ttu-id="e91c6-158">set()</span><span class="sxs-lookup"><span data-stu-id="e91c6-158">set()</span></span>
- <span data-ttu-id="e91c6-159">setDirty() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="e91c6-159">setDirty() (preview)</span></span>
- <span data-ttu-id="e91c6-160">toJSON()</span><span class="sxs-lookup"><span data-stu-id="e91c6-160">toJSON()</span></span>
- <span data-ttu-id="e91c6-161">track()</span><span class="sxs-lookup"><span data-stu-id="e91c6-161">track()</span></span>
- <span data-ttu-id="e91c6-162">untrack()</span><span class="sxs-lookup"><span data-stu-id="e91c6-162">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="e91c6-163">Méthodes et propriétés propres à une langue RangeArea</span><span class="sxs-lookup"><span data-stu-id="e91c6-163">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="e91c6-164">Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet.</span><span class="sxs-lookup"><span data-stu-id="e91c6-164">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="e91c6-165">Ce qui est une sélection de certains d’entre eux :</span><span class="sxs-lookup"><span data-stu-id="e91c6-165">The following is a selection of them:</span></span>

- <span data-ttu-id="e91c6-166">`areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet.</span><span class="sxs-lookup"><span data-stu-id="e91c6-166">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="e91c6-167">L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel.</span><span class="sxs-lookup"><span data-stu-id="e91c6-167">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="e91c6-168">Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="e91c6-168">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="e91c6-169">`areaCount`: Le nombre total de plages dans le`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-169">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e91c6-170">`getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-170">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="e91c6-171">Créer l’objet RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e91c6-171">Create RangeAreas</span></span>

<span data-ttu-id="e91c6-172">Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:</span><span class="sxs-lookup"><span data-stu-id="e91c6-172">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="e91c6-173">Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="e91c6-173">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="e91c6-174">Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="e91c6-174">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="e91c6-175">Appel `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-175">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="e91c6-176">Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.</span><span class="sxs-lookup"><span data-stu-id="e91c6-176">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="e91c6-177">Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-177">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="e91c6-178">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-178">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="e91c6-179">Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-179">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="e91c6-180">N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-180">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="e91c6-181">Cela mènera à un comportement indésirable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="e91c6-181">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="e91c6-182">Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas.</span><span class="sxs-lookup"><span data-stu-id="e91c6-182">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="e91c6-183">Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-183">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="e91c6-184">De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet*soit*supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore.</span><span class="sxs-lookup"><span data-stu-id="e91c6-184">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="e91c6-185">Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.</span><span class="sxs-lookup"><span data-stu-id="e91c6-185">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="e91c6-186">Définir les propriétés sur plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="e91c6-186">Set properties on multiple ranges</span></span>

<span data-ttu-id="e91c6-187">Paramétrer une propriété sur un objet `RangeAreas` établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-187">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="e91c6-188">Ce qui suit est un exemple de paramétrage d’une propriété sur des plages multiples.</span><span class="sxs-lookup"><span data-stu-id="e91c6-188">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="e91c6-189">La fonction surligne les plages**F3:F5** and **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="e91c6-189">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="e91c6-190">Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution.</span><span class="sxs-lookup"><span data-stu-id="e91c6-190">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="e91c6-191">Certains des scénarios dans lesquels ceci peut s’appliquer incluent:</span><span class="sxs-lookup"><span data-stu-id="e91c6-191">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="e91c6-192">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="e91c6-192">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="e91c6-193">Le code s’exécute dans le contexte de données importées où le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="e91c6-193">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="e91c6-194">Obtenir des cellules spéciales à partir de plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="e91c6-194">Get special cells from multiple ranges</span></span>

<span data-ttu-id="e91c6-195">Les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` sur l’objet `RangeAreas` fonctionnent de manière analogue aux méthodes du même nom sur l’objet `Range`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-195">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="e91c6-196">Ces méthodes retournent les cellules disposant de la caractéristique spécifiée à partir de toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-196">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="e91c6-197">Voir la section [Trouver des cellules spéciales dans une plage](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) pour plus d’informations sur les cellules spéciales.</span><span class="sxs-lookup"><span data-stu-id="e91c6-197">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) section for more details on special cells.</span></span>

<span data-ttu-id="e91c6-198">Lors de l’appel de la méthode `getSpecialCells` ou `getSpecialCellsOrNullObject` sur un objet `RangeAreas` :</span><span class="sxs-lookup"><span data-stu-id="e91c6-198">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="e91c6-199">si vous passez `Excel.SpecialCellType.sameConditionalFormat` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même mise en forme conditionnelle que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-199">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="e91c6-200">Si vous passez `Excel.SpecialCellType.sameDataValidation` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même règle de validation des données que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-200">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="e91c6-201">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e91c6-201">Read properties of RangeAreas</span></span>

<span data-ttu-id="e91c6-202">La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-202">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="e91c6-203">La règle générales est que si une valeur consistante*peut*être renvoyée, elle sera renvoyée.</span><span class="sxs-lookup"><span data-stu-id="e91c6-203">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="e91c6-204">Par exemple, dans le code suivant, le code RGB pour rose (`#FFC0CB`) et`true`sera connecté à la console car les deux plages dans l’objet`RangeAreas` dispose d’un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="e91c6-204">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="e91c6-205">Les choses se compliquent lorsque la consistance est impossible.</span><span class="sxs-lookup"><span data-stu-id="e91c6-205">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="e91c6-206">Le comportement de propriétés`RangeAreas` suit ces trois principes:</span><span class="sxs-lookup"><span data-stu-id="e91c6-206">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="e91c6-207">Une propriété booléenne d’un objet`RangeAreas` renvoie`false`à moins que la propriété soit vraie pour toutes les plages membres.</span><span class="sxs-lookup"><span data-stu-id="e91c6-207">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="e91c6-208">Les propriétés non-booléennes, avec l’exception de la propriété`address`renvoie`null`à moins que la propriété correspondante sur toutes les plages membres dispose de la même valeur.</span><span class="sxs-lookup"><span data-stu-id="e91c6-208">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="e91c6-209">La propriété`address`renvoie une chaîne délimitée à virgule des adresses des plages membres.</span><span class="sxs-lookup"><span data-stu-id="e91c6-209">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="e91c6-210">Par exemple, le code suivante crée un`RangeAreas`dans lequel seule une plage est une colonne entière et seule une est remplie de rose.</span><span class="sxs-lookup"><span data-stu-id="e91c6-210">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="e91c6-211">La console s’affichera`null`pour un remplissage de couleur,`false`pour la propriété`isEntireRow` et «Sheet1!F3:F5, Sheet1!H:H» (en présumant que la feuille de calcule soit «Sheet1») pour la propriété`address`.</span><span class="sxs-lookup"><span data-stu-id="e91c6-211">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e91c6-212">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e91c6-212">See also</span></span>

- [<span data-ttu-id="e91c6-213">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e91c6-213">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="e91c6-214">Utilisation de plages à l’aide de l’API JavaScript pour Excel (basique)</span><span class="sxs-lookup"><span data-stu-id="e91c6-214">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="e91c6-215">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="e91c6-215">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
