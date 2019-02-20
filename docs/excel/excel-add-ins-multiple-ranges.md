---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: ''
ms.date: 12/26/2018
localization_priority: Normal
ms.openlocfilehash: ba171026f410e59e81d3bbb6bc0f799c929a25a3
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/19/2019
ms.locfileid: "29387869"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="ccb11-102">Travailler simultanément avec plusieurs plages dans des compléments Excel (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="ccb11-102">Work with multiple ranges simultaneously in Excel add-ins (preview)</span></span>

<span data-ttu-id="ccb11-103">La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages.</span><span class="sxs-lookup"><span data-stu-id="ccb11-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="ccb11-104">Les plages n’ont pas besoin d’être contigus.</span><span class="sxs-lookup"><span data-stu-id="ccb11-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="ccb11-105">En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.</span><span class="sxs-lookup"><span data-stu-id="ccb11-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="ccb11-106">Les APIs décrits dans cet article nécessitent**la version Office 2016 «Démarrer en un Clic» 1809 Build 10820.20000**ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ccb11-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="ccb11-107">(Vous devrez peut-être rejoindre le[programme Office Insider](https://products.office.com/office-insider) pour obtenir une build appropriée.) Par ailleurs, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir de [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="ccb11-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="ccb11-108">Enfin, nous n’avons pas encore les pages de référence pour ces APIs.</span><span class="sxs-lookup"><span data-stu-id="ccb11-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="ccb11-109">Mais le fichier de type définition suivant comporte des descriptions à leur place: [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="ccb11-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="ccb11-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="ccb11-110">RangeAreas</span></span>

<span data-ttu-id="ccb11-111">Un ensemble de plages (voire non contiguës) est représenté par un `Excel.RangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="ccb11-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="ccb11-112">Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:</span><span class="sxs-lookup"><span data-stu-id="ccb11-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="ccb11-113">Les types de données pour les propriétés et le comportement des méthodes et des getters.</span><span class="sxs-lookup"><span data-stu-id="ccb11-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="ccb11-114">Les types de données de paramètres et des comportements de la méthode.</span><span class="sxs-lookup"><span data-stu-id="ccb11-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="ccb11-115">Les types de données de méthodes renvoient des valeurs.</span><span class="sxs-lookup"><span data-stu-id="ccb11-115">The data types of method return values.</span></span>

<span data-ttu-id="ccb11-116">Quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="ccb11-116">Some examples:</span></span>

- <span data-ttu-id="ccb11-117">`RangeAreas` a une`address` propriété qui renvoie une chaîne séparée par des virgules de plage d’adresses, au lieu d’une adresse comme avec la `Range.address` propriété.</span><span class="sxs-lookup"><span data-stu-id="ccb11-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="ccb11-118">`RangeAreas` a une`dataValidation` propriété qui renvoie un`DataValidation` objet qui représente la validation des données de toutes les plages dans la `RangeAreas`, s’il est cohérent.</span><span class="sxs-lookup"><span data-stu-id="ccb11-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="ccb11-119">La propriété est`null` si les objets`DataValidation` identiques ne sont pas appliqués à toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="ccb11-120">Il s’agit d’un principe général, mais pas universel avec l’`RangeAreas` objet: *si une propriété ne comporte pas de valeurs cohérentes sur tous les plages dans la`RangeAreas`, cela signifie`null`.*</span><span class="sxs-lookup"><span data-stu-id="ccb11-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="ccb11-121">Voir[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) pour plus d’informations et quelques exceptions.</span><span class="sxs-lookup"><span data-stu-id="ccb11-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="ccb11-122">`RangeAreas.cellCount` Obtient le nombre total de cellules dans toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="ccb11-123">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="ccb11-124">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` retourner un autre`RangeAreas` objet qui représente toutes les colonnes (ou lignes) dans toutes les plages dans la `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="ccb11-125">Par exemple, si le`RangeAreas` représente « A1 : C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un`RangeAreas` objet qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="ccb11-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="ccb11-126">`RangeAreas.copyFrom` peut prendre soit un`Range` ou d’un`RangeAreas` paramètre représentant la ou les plage(s) source de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="ccb11-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="ccb11-127">La liste complète des membres plage sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="ccb11-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="ccb11-128">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ccb11-128">Properties</span></span>

<span data-ttu-id="ccb11-129">Être familiarisé avec[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) avant d’écrire de code qui lit les propriétés répertoriées.</span><span class="sxs-lookup"><span data-stu-id="ccb11-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="ccb11-130">Il existe des subtilités sur ce qui est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="ccb11-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="ccb11-131">address</span><span class="sxs-lookup"><span data-stu-id="ccb11-131">address</span></span>
- <span data-ttu-id="ccb11-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="ccb11-132">addressLocal</span></span>
- <span data-ttu-id="ccb11-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="ccb11-133">cellCount</span></span>
- <span data-ttu-id="ccb11-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="ccb11-134">conditionalFormats</span></span>
- <span data-ttu-id="ccb11-135">context</span><span class="sxs-lookup"><span data-stu-id="ccb11-135">context</span></span>
- <span data-ttu-id="ccb11-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="ccb11-136">dataValidation</span></span>
- <span data-ttu-id="ccb11-137">format</span><span class="sxs-lookup"><span data-stu-id="ccb11-137">format</span></span>
- <span data-ttu-id="ccb11-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="ccb11-138">isEntireColumn</span></span>
- <span data-ttu-id="ccb11-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="ccb11-139">isEntireRow</span></span>
- <span data-ttu-id="ccb11-140">style</span><span class="sxs-lookup"><span data-stu-id="ccb11-140">style</span></span>
- <span data-ttu-id="ccb11-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="ccb11-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="ccb11-142">Méthodes</span><span class="sxs-lookup"><span data-stu-id="ccb11-142">Methods</span></span>

<span data-ttu-id="ccb11-143">Les méthodes de plage dans l’aperçu sont marquées.</span><span class="sxs-lookup"><span data-stu-id="ccb11-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="ccb11-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="ccb11-144">calculate()</span></span>
- <span data-ttu-id="ccb11-145">clear()</span><span class="sxs-lookup"><span data-stu-id="ccb11-145">clear()</span></span>
- <span data-ttu-id="ccb11-146">convertDataTypeToText() (preview)</span><span class="sxs-lookup"><span data-stu-id="ccb11-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="ccb11-147">convertToLinkedDataType() (preview)</span><span class="sxs-lookup"><span data-stu-id="ccb11-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="ccb11-148">copyFrom() (preview)</span><span class="sxs-lookup"><span data-stu-id="ccb11-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="ccb11-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="ccb11-149">getEntireColumn()</span></span>
- <span data-ttu-id="ccb11-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="ccb11-150">getEntireRow()</span></span>
- <span data-ttu-id="ccb11-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="ccb11-151">getIntersection()</span></span>
- <span data-ttu-id="ccb11-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="ccb11-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="ccb11-153">getOffsetRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="ccb11-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="ccb11-154">getSpecialCells() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="ccb11-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="ccb11-155">getSpecialCellsOrNullObject() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="ccb11-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="ccb11-156">getTables() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="ccb11-156">getTables() (preview)</span></span>
- <span data-ttu-id="ccb11-157">getUsedRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="ccb11-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="ccb11-158">getUsedRangeOrNullObject() (appelé getUsedRangeAreasOrNullObject sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="ccb11-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="ccb11-159">load()</span><span class="sxs-lookup"><span data-stu-id="ccb11-159">load()</span></span>
- <span data-ttu-id="ccb11-160">set()</span><span class="sxs-lookup"><span data-stu-id="ccb11-160">set()</span></span>
- <span data-ttu-id="ccb11-161">setDirty() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="ccb11-161">setDirty() (preview)</span></span>
- <span data-ttu-id="ccb11-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="ccb11-162">toJSON()</span></span>
- <span data-ttu-id="ccb11-163">track()</span><span class="sxs-lookup"><span data-stu-id="ccb11-163">track()</span></span>
- <span data-ttu-id="ccb11-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="ccb11-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="ccb11-165">Méthodes et propriétés propres à une langue RangeArea</span><span class="sxs-lookup"><span data-stu-id="ccb11-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="ccb11-166">Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet.</span><span class="sxs-lookup"><span data-stu-id="ccb11-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="ccb11-167">Ce qui est une sélection de certains d’entre eux :</span><span class="sxs-lookup"><span data-stu-id="ccb11-167">The following is a selection of them:</span></span>

- <span data-ttu-id="ccb11-168">`areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet.</span><span class="sxs-lookup"><span data-stu-id="ccb11-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="ccb11-169">L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel.</span><span class="sxs-lookup"><span data-stu-id="ccb11-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="ccb11-170">Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="ccb11-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="ccb11-171">`areaCount`: Le nombre total de plages dans le`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="ccb11-172">`getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="ccb11-173">Créer l’objet RangeAreas</span><span class="sxs-lookup"><span data-stu-id="ccb11-173">Create RangeAreas</span></span>

<span data-ttu-id="ccb11-174">Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:</span><span class="sxs-lookup"><span data-stu-id="ccb11-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="ccb11-175">Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="ccb11-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="ccb11-176">Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="ccb11-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="ccb11-177">Appel `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="ccb11-178">Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.</span><span class="sxs-lookup"><span data-stu-id="ccb11-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="ccb11-179">Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="ccb11-180">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="ccb11-181">Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="ccb11-182">N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="ccb11-183">Cela mènera à un comportement indésirable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="ccb11-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="ccb11-184">Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas.</span><span class="sxs-lookup"><span data-stu-id="ccb11-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="ccb11-185">Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="ccb11-186">De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet*soit*supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore.</span><span class="sxs-lookup"><span data-stu-id="ccb11-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="ccb11-187">Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.</span><span class="sxs-lookup"><span data-stu-id="ccb11-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="ccb11-188">Définir les propriétés sur plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="ccb11-188">Set properties on multiple ranges</span></span>

<span data-ttu-id="ccb11-189">Paramétrer une propriété sur un objet `RangeAreas` établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-189">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="ccb11-190">Ce qui suit est un exemple de paramétrage d’une propriété sur des plages multiples.</span><span class="sxs-lookup"><span data-stu-id="ccb11-190">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="ccb11-191">La fonction surligne les plages**F3:F5** and **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="ccb11-191">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="ccb11-192">Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution.</span><span class="sxs-lookup"><span data-stu-id="ccb11-192">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="ccb11-193">Certains des scénarios dans lesquels ceci peut s’appliquer incluent:</span><span class="sxs-lookup"><span data-stu-id="ccb11-193">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="ccb11-194">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="ccb11-194">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="ccb11-195">Le code s’exécute dans le contexte de données importées où le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="ccb11-195">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="ccb11-196">Obtenir des cellules spéciales à partir de plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="ccb11-196">Get special cells from multiple ranges</span></span>

<span data-ttu-id="ccb11-197">Les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` sur l’objet `RangeAreas` fonctionnent de manière analogue aux méthodes du même nom sur l’objet `Range`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-197">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="ccb11-198">Ces méthodes retournent les cellules disposant de la caractéristique spécifiée à partir de toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-198">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="ccb11-199">Voir la section [Trouver des cellules spéciales dans une plage](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) pour plus d’informations sur les cellules spéciales.</span><span class="sxs-lookup"><span data-stu-id="ccb11-199">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) section for more details on special cells.</span></span>

<span data-ttu-id="ccb11-200">Lors de l’appel de la méthode `getSpecialCells` ou `getSpecialCellsOrNullObject` sur un objet `RangeAreas` :</span><span class="sxs-lookup"><span data-stu-id="ccb11-200">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="ccb11-201">si vous passez `Excel.SpecialCellType.sameConditionalFormat` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même mise en forme conditionnelle que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-201">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="ccb11-202">Si vous passez `Excel.SpecialCellType.sameDataValidation` en tant que premier paramètre, la méthode renvoie toutes les cellules disposant de la même règle de validation des données que la cellule supérieure gauche de la première plage dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-202">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="ccb11-203">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="ccb11-203">Read properties of RangeAreas</span></span>

<span data-ttu-id="ccb11-204">La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-204">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="ccb11-205">La règle générales est que si une valeur consistante*peut*être renvoyée, elle sera renvoyée.</span><span class="sxs-lookup"><span data-stu-id="ccb11-205">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="ccb11-206">Par exemple, dans le code suivant, le code RGB pour rose (`#FFC0CB`) et`true`sera connecté à la console car les deux plages dans l’objet`RangeAreas` dispose d’un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="ccb11-206">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="ccb11-207">Les choses se compliquent lorsque la consistance est impossible.</span><span class="sxs-lookup"><span data-stu-id="ccb11-207">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="ccb11-208">Le comportement de propriétés`RangeAreas` suit ces trois principes:</span><span class="sxs-lookup"><span data-stu-id="ccb11-208">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="ccb11-209">Une propriété booléenne d’un objet`RangeAreas` renvoie`false`à moins que la propriété soit vraie pour toutes les plages membres.</span><span class="sxs-lookup"><span data-stu-id="ccb11-209">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="ccb11-210">Les propriétés non-booléennes, avec l’exception de la propriété`address`renvoie`null`à moins que la propriété correspondante sur toutes les plages membres dispose de la même valeur.</span><span class="sxs-lookup"><span data-stu-id="ccb11-210">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="ccb11-211">La propriété`address`renvoie une chaîne délimitée à virgule des adresses des plages membres.</span><span class="sxs-lookup"><span data-stu-id="ccb11-211">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="ccb11-212">Par exemple, le code suivante crée un`RangeAreas`dans lequel seule une plage est une colonne entière et seule une est remplie de rose.</span><span class="sxs-lookup"><span data-stu-id="ccb11-212">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="ccb11-213">La console s’affichera`null`pour un remplissage de couleur,`false`pour la propriété`isEntireRow` et «Sheet1!F3:F5, Sheet1!H:H» (en présumant que la feuille de calcule soit «Sheet1») pour la propriété`address`.</span><span class="sxs-lookup"><span data-stu-id="ccb11-213">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="ccb11-214">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ccb11-214">See also</span></span>

- [<span data-ttu-id="ccb11-215">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ccb11-215">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="ccb11-216">Utilisation de plages à l’aide de l’API JavaScript pour Excel (basique)</span><span class="sxs-lookup"><span data-stu-id="ccb11-216">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="ccb11-217">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="ccb11-217">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
