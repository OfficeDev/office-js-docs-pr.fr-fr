---
title: Travailler simultanément avec plusieurs plages dans des compléments Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: f1217fc76d14269882a73ec5eb7758e519563456
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383224"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="b8045-102">Travailler simultanément avec plusieurs plages dans des compléments Excel (Prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="b8045-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="b8045-103">La bibliothèque JavaScript Excel permet à votre complément d’effectuer des opérations et définir des propriétés, de manière simultanée sur plusieurs plages.</span><span class="sxs-lookup"><span data-stu-id="b8045-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="b8045-104">Les plages n’ont pas besoin d’être contigus.</span><span class="sxs-lookup"><span data-stu-id="b8045-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="b8045-105">En plus de rendre votre code plus simple, cette manière de paramétrer une propriété s’exécute beaucoup plus rapidement que paramétrer la même propriété pour chaque les plages de manière individuelle.</span><span class="sxs-lookup"><span data-stu-id="b8045-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="b8045-106">Les APIs décrits dans cet article nécessitent**la version Office 2016 «Démarrer en un Clic» 1809 Build 10820.20000**ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b8045-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="b8045-107">(Vous devrez peut-être rejoindre le[programme Office Insider](https://products.office.com/office-insider) pour obtenir une build appropriée.) Par ailleurs, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir de [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="b8045-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="b8045-108">Enfin, nous n’avons pas encore les pages de référence pour ces APIs.</span><span class="sxs-lookup"><span data-stu-id="b8045-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="b8045-109">Mais le fichier de type définition suivant comporte des descriptions à leur place: [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="b8045-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="b8045-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="b8045-110">RangeAreas</span></span>

<span data-ttu-id="b8045-111">Un ensemble de plages (voire non contiguës) est représenté par un `Excel.RangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="b8045-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="b8045-112">Il possède des propriétés et des méthodes similaires au type`Range` (bon nombre des noms identiques ou similaires,), mais les ajustements ont été apportées à:</span><span class="sxs-lookup"><span data-stu-id="b8045-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="b8045-113">Les types de données pour les propriétés et le comportement des méthodes et des getters.</span><span class="sxs-lookup"><span data-stu-id="b8045-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="b8045-114">Les types de données de paramètres et des comportements de la méthode.</span><span class="sxs-lookup"><span data-stu-id="b8045-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="b8045-115">Les types de données de méthodes renvoient des valeurs.</span><span class="sxs-lookup"><span data-stu-id="b8045-115">The data types of method return values.</span></span>

<span data-ttu-id="b8045-116">Quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="b8045-116">Some examples:</span></span>

- <span data-ttu-id="b8045-117">`RangeAreas` a une`address` propriété qui renvoie une chaîne séparée par des virgules de plage d’adresses, au lieu d’une adresse comme avec la `Range.address` propriété.</span><span class="sxs-lookup"><span data-stu-id="b8045-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="b8045-118">`RangeAreas` a une`dataValidation` propriété qui renvoie un`DataValidation` objet qui représente la validation des données de toutes les plages dans la `RangeAreas`, s’il est cohérent.</span><span class="sxs-lookup"><span data-stu-id="b8045-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="b8045-119">La propriété est`null` si les objets`DataValidation` identiques ne sont pas appliqués à toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="b8045-120">Il s’agit d’un principe général, mais pas universel avec l’`RangeAreas` objet: *si une propriété ne comporte pas de valeurs cohérentes sur tous les plages dans la`RangeAreas`, cela signifie`null`.*</span><span class="sxs-lookup"><span data-stu-id="b8045-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="b8045-121">Voir[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) pour plus d’informations et quelques exceptions.</span><span class="sxs-lookup"><span data-stu-id="b8045-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="b8045-122">`RangeAreas.cellCount` Obtient le nombre total de cellules dans toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="b8045-123">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans la`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="b8045-124">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` retourner un autre`RangeAreas` objet qui représente toutes les colonnes (ou lignes) dans toutes les plages dans la `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="b8045-125">Par exemple, si le`RangeAreas` représente « A1 : C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un`RangeAreas` objet qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="b8045-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="b8045-126">`RangeAreas.copyFrom` peut prendre soit un`Range` ou d’un`RangeAreas` paramètre représentant la ou les plage(s) source de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="b8045-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="b8045-127">La liste complète des membres plage sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="b8045-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="b8045-128">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b8045-128">Properties</span></span>

<span data-ttu-id="b8045-129">Être familiarisé avec[Lire les propriétés de RangeAreas](#read-properties-of-rangeareas) avant d’écrire de code qui lit les propriétés répertoriées.</span><span class="sxs-lookup"><span data-stu-id="b8045-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="b8045-130">Il existe des subtilités sur ce qui est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="b8045-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="b8045-131">address</span><span class="sxs-lookup"><span data-stu-id="b8045-131">address</span></span>
- <span data-ttu-id="b8045-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="b8045-132">addressLocal</span></span>
- <span data-ttu-id="b8045-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="b8045-133">cellCount</span></span>
- <span data-ttu-id="b8045-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="b8045-134">conditionalFormats</span></span>
- <span data-ttu-id="b8045-135">context</span><span class="sxs-lookup"><span data-stu-id="b8045-135">context</span></span>
- <span data-ttu-id="b8045-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="b8045-136">dataValidation</span></span>
- <span data-ttu-id="b8045-137">format</span><span class="sxs-lookup"><span data-stu-id="b8045-137">format</span></span>
- <span data-ttu-id="b8045-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="b8045-138">isEntireColumn</span></span>
- <span data-ttu-id="b8045-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="b8045-139">isEntireRow</span></span>
- <span data-ttu-id="b8045-140">style</span><span class="sxs-lookup"><span data-stu-id="b8045-140">style</span></span>
- <span data-ttu-id="b8045-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="b8045-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="b8045-142">Méthodes</span><span class="sxs-lookup"><span data-stu-id="b8045-142">Methods</span></span>

<span data-ttu-id="b8045-143">Les méthodes de plage dans l’aperçu sont marquées.</span><span class="sxs-lookup"><span data-stu-id="b8045-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="b8045-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="b8045-144">calculate()</span></span>
- <span data-ttu-id="b8045-145">clear()</span><span class="sxs-lookup"><span data-stu-id="b8045-145">clear()</span></span>
- <span data-ttu-id="b8045-146">convertDataTypeToText() (preview)</span><span class="sxs-lookup"><span data-stu-id="b8045-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="b8045-147">convertToLinkedDataType() (preview)</span><span class="sxs-lookup"><span data-stu-id="b8045-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="b8045-148">copyFrom() (preview)</span><span class="sxs-lookup"><span data-stu-id="b8045-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="b8045-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="b8045-149">getEntireColumn()</span></span>
- <span data-ttu-id="b8045-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="b8045-150">getEntireRow()</span></span>
- <span data-ttu-id="b8045-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="b8045-151">getIntersection()</span></span>
- <span data-ttu-id="b8045-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="b8045-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="b8045-153">getOffsetRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="b8045-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="b8045-154">getSpecialCells() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="b8045-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="b8045-155">getSpecialCellsOrNullObject() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="b8045-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="b8045-156">getTables() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="b8045-156">getTables() (preview)</span></span>
- <span data-ttu-id="b8045-157">getUsedRange() (appelé getOffsetRangeAreas sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="b8045-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="b8045-158">getUsedRangeOrNullObject() (appelé getUsedRangeAreasOrNullObject sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="b8045-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="b8045-159">load()</span><span class="sxs-lookup"><span data-stu-id="b8045-159">load()</span></span>
- <span data-ttu-id="b8045-160">set()</span><span class="sxs-lookup"><span data-stu-id="b8045-160">set()</span></span>
- <span data-ttu-id="b8045-161">setDirty() (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="b8045-161">setDirty() (preview)</span></span>
- <span data-ttu-id="b8045-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="b8045-162">toJSON()</span></span>
- <span data-ttu-id="b8045-163">track()</span><span class="sxs-lookup"><span data-stu-id="b8045-163">track()</span></span>
- <span data-ttu-id="b8045-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="b8045-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="b8045-165">Méthodes et propriétés propres à une langue RangeArea</span><span class="sxs-lookup"><span data-stu-id="b8045-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="b8045-166">Le `RangeAreas` type possède des propriétés et des méthodes qui ne sont pas sur l’`Range`objet.</span><span class="sxs-lookup"><span data-stu-id="b8045-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="b8045-167">Ce qui est une sélection de certains d’entre eux :</span><span class="sxs-lookup"><span data-stu-id="b8045-167">The following is a selection of them:</span></span>

- <span data-ttu-id="b8045-168">`areas`: A`RangeCollection` objet qui contient toutes les plages représentées par l’ `RangeAreas`objet.</span><span class="sxs-lookup"><span data-stu-id="b8045-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="b8045-169">L’`RangeCollection`objet est également nouveau et est semblable à d’autres objets de collection de sites Excel.</span><span class="sxs-lookup"><span data-stu-id="b8045-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="b8045-170">Il possède une`items`propriété est une matrice d’`Range` objets représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="b8045-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="b8045-171">`areaCount`: Le nombre total de plages dans le`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="b8045-172">`getOffsetRangeAreas`: Fonctionne comme[Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’une `RangeAreas` est renvoyée et il contient des plages sont en décalage avec des plages du fichier d’origine`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="b8045-173">Créer RangeAreas et définir les propriétés</span><span class="sxs-lookup"><span data-stu-id="b8045-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="b8045-174">Vous pouvez créer`RangeAreas`l’objet selon deux méthodes de base:</span><span class="sxs-lookup"><span data-stu-id="b8045-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="b8045-175">Appeler`Worksheet.getRanges()` et de transmettre une chaîne comportant des adresses de plage séparées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="b8045-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="b8045-176">Si une plage que vous souhaitez inclure a été modifiée en[NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="b8045-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="b8045-177">Appel `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="b8045-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="b8045-178">Cette méthode renvoie une`RangeAreas`représentation de toutes les plages qui sont sélectionnées sur la feuille de calcul active actuelle.</span><span class="sxs-lookup"><span data-stu-id="b8045-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="b8045-179">Une fois que vous avez un`RangeAreas`objet, vous pouvez en créer d’autres à l’aide des méthodes sur l’objet qui renvoie`RangeAreas`tel que`getOffsetRangeAreas`et`getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="b8045-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="b8045-180">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="b8045-181">Par exemple, la collection dans`RangeAreas.areas`n’a pas une méthode`add`.</span><span class="sxs-lookup"><span data-stu-id="b8045-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="b8045-182">N’essayez pas d’ajouter ou de supprimer directement les membres du tableau`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="b8045-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="b8045-183">Cela mènera à un comportement indésirable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="b8045-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="b8045-184">Par exemple, il est possible de pousser un objet`Range` supplémentaire sur le tableau, mais ceci entrainera des erreurs car les propriétés`RangeAreas`et les méthodes se comportent comme si le nouvel élément n’existait pas.</span><span class="sxs-lookup"><span data-stu-id="b8045-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="b8045-185">Par exemple, la propriété`areaCount`n’inclut pas les plages poussées de cette manière, et le `RangeAreas.getItemAt(index)` lance une erreur si`index`est plus grand que`areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="b8045-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="b8045-186">De la même façon, supprimer un objet`Range`dans la plage`RangeAreas.areas.items`en obtenant une référence liée à celui-ci et en appelant sa méthode`Range.delete` entraîne des bogues: bien que `Range`l’objet*soit*supprimé, les propriétés et les méthodes de l’objet`RangeAreas`parent se comporte, ou tente de le faire, comme s’il existait encore.</span><span class="sxs-lookup"><span data-stu-id="b8045-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="b8045-187">Par exemple, si votre code appelle`RangeAreas.calculate`, Office essaiera de calculer la plage, mais comprendra une erreur car l’objet de la plage n’est plus là.</span><span class="sxs-lookup"><span data-stu-id="b8045-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="b8045-188">Paramétrer une propriété sur un`RangeAreas`établit une propriété correspondante sur toutes les plages dans la collection`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="b8045-189">Le suivant est un exemple de paramétrage d’une propriété sur des plages multiples.</span><span class="sxs-lookup"><span data-stu-id="b8045-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="b8045-190">La fonction surligne les plages**F3:F5** and **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="b8045-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="b8045-191">Cet exemple s’applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de plage que vous passez à`getRanges`ou facilement les calculer à l’exécution.</span><span class="sxs-lookup"><span data-stu-id="b8045-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="b8045-192">Certains des scénarios dans lesquels ceci peut s’appliquer incluent:</span><span class="sxs-lookup"><span data-stu-id="b8045-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="b8045-193">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="b8045-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="b8045-194">Le code s’exécute dans le contexte de données importées où le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="b8045-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="b8045-195">Lorsque vous ne pouvez pas connaitre au moment de coder quelles plages sont nécessaires pour opérer, vous devez les découvrir lors de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="b8045-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="b8045-196">La prochaine section traite de ces scénarios.</span><span class="sxs-lookup"><span data-stu-id="b8045-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="b8045-197">Découvrez les zones de plage au niveau de la programmation</span><span class="sxs-lookup"><span data-stu-id="b8045-197">Discover range areas programmatically</span></span>

<span data-ttu-id="b8045-198">Les méthodes `Range.getSpecialCells()`et`Range.getSpecialCellsOrNullObject()`vous permettent de trouver lors de l’exécution les plages que vous souhaitez faire fonctionner sur la base des caractéristiques des cellules et du type des valeurs des cellules.</span><span class="sxs-lookup"><span data-stu-id="b8045-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="b8045-199">Voici les signatures des méthodes à partir des types de fichiers de données TypeScript:</span><span class="sxs-lookup"><span data-stu-id="b8045-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="b8045-200">Voici un exemple d’utilisation de la première.</span><span class="sxs-lookup"><span data-stu-id="b8045-200">The following is an example of using the first one.</span></span> <span data-ttu-id="b8045-201">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="b8045-201">About this code, note:</span></span>

- <span data-ttu-id="b8045-202">Cela limite la partie de la feuille qui nécessite d’être recherchée en appelant d’abord`Worksheet.getUsedRange`et en appelant`getSpecialCells`uniquement pour cette plage.</span><span class="sxs-lookup"><span data-stu-id="b8045-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="b8045-203">Il passe comme paramètre à la version chaîne`getSpecialCells`d’une valeur à partir du enum`Excel.SpecialCellType`.</span><span class="sxs-lookup"><span data-stu-id="b8045-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="b8045-204">Certaines des autres valeurs qui peuvent être passées à la place sont «Vides»pour des cellules vides, «Constantes» pour des cellules avec des valeurs littérales au lieu des formules, et «SameConditionalFormat» pour les cellules qui disposent de la même mise en forme conditionnelle comme la première cellule dans le`usedRange`.</span><span class="sxs-lookup"><span data-stu-id="b8045-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="b8045-205">La première cellule est la cellule en haut toute à gauche.</span><span class="sxs-lookup"><span data-stu-id="b8045-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="b8045-206">Pour une liste complète des valeurs dans l’enum, voir[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="b8045-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="b8045-207">La`getSpecialCells`méthode renvoie un`RangeAreas`objet, toutes les cellules alors dotées de formules seront colorées en rose même si elles ne sont pas adjacentes.</span><span class="sxs-lookup"><span data-stu-id="b8045-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="b8045-208">Parfois la plage ne dispose pas *de*cellules avec la caractéristique ciblée.</span><span class="sxs-lookup"><span data-stu-id="b8045-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="b8045-209">Si`getSpecialCells`n’en trouve pas, elle lance une erreur**ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="b8045-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="b8045-210">Cela dévie le flux de contrôle vers un(e)`catch`bloc/méthode, s’il en existe.</span><span class="sxs-lookup"><span data-stu-id="b8045-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="b8045-211">S’il n’existe pas, l’erreur arrête la fonction.</span><span class="sxs-lookup"><span data-stu-id="b8045-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="b8045-212">Il peut avoir des scénarios dans lesquels émettre l’erreur est exactement ce que vous souhaitez lorsqu’il n’y a pas de cellules avec de caractéristique ciblée.</span><span class="sxs-lookup"><span data-stu-id="b8045-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="b8045-213">Mais dans les scénarios dans lesquels cela est normal, mais peut-être gênant, pour les cellules qui correspondent pas; votre code doit vérifier cette possibilité et le gérer gracieusement sans émettre d’erreur.</span><span class="sxs-lookup"><span data-stu-id="b8045-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="b8045-214">Pour ces scénarios, utilisez la méthode`getSpecialCellsOrNullObject` et testez la propriété`RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="b8045-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="b8045-215">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="b8045-215">The following is an example.</span></span> <span data-ttu-id="b8045-216">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="b8045-216">Note about this code:</span></span>

- <span data-ttu-id="b8045-217">La méthode`getSpecialCellsOrNullObject`renvoie toujours un objet proxy, donc il ne s’agit jamais du sens`null`JavaScript ordinaire.</span><span class="sxs-lookup"><span data-stu-id="b8045-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="b8045-218">Mais si les cellules non correspondantes sont introuvables, la propriété`isNullObject` de l’objet est établi à`true`.</span><span class="sxs-lookup"><span data-stu-id="b8045-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="b8045-219">Il appelle`context.sync`*avant*de tester la propriété`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="b8045-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="b8045-220">Il s’agit d’une condition avec toutes les méthodes et propriétés`*OrNullObject`, car vous devez toujours télécharger et synchroniser une propriété afin de le lire. </span><span class="sxs-lookup"><span data-stu-id="b8045-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="b8045-221">Cependant, il n’est pas nécessaire de télécharger*de manière explicite*la propriété`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="b8045-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="b8045-222">Il est automatiquement téléchargé par le`context.sync`même si`load`n’est pas appelé sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="b8045-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="b8045-223">Pour plus d'informations, consultez le[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="b8045-223">For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="b8045-224">Vous pouvez tester ce code en sélectionnant d’abord une plage qui n’a pas de cellules de formule et en l’exécutant.</span><span class="sxs-lookup"><span data-stu-id="b8045-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="b8045-225">Puis sélectionnez une plage qui dispose au moins d’une cellule dotée d’une formule et en l’exécutant à nouveau.</span><span class="sxs-lookup"><span data-stu-id="b8045-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="b8045-226">Par simplicité, tous les autres exemples dans cet article, utilisez la méthode`getSpecialCells`au lieu de`getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="b8045-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="b8045-227">Réduisez les cellules cibles avec les types de valeur de cellule</span><span class="sxs-lookup"><span data-stu-id="b8045-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="b8045-228">Il existe un paramètre secondaire optionnel, de type enum`Excel.SpecialCellValueType`, qui réduise encore la cellule à la cible.</span><span class="sxs-lookup"><span data-stu-id="b8045-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="b8045-229">Vous pouvez l’utiliser uniquement lorsque vous passez soit «Formules» ou «Constantes» à`getSpecialCells`ou`getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="b8045-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="b8045-230">Le paramètre spécifie que vous souhaitez uniquement les cellules avec certains types de valeurs.</span><span class="sxs-lookup"><span data-stu-id="b8045-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="b8045-231">Il existe quatre types de base: «Erreur», «Logique»(ce qui signifie booléen), «Nombres», et «Texte».</span><span class="sxs-lookup"><span data-stu-id="b8045-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="b8045-232">(L’enum dispose d’autres valeurs hormis les quatre traités plus haut.) Ce qui suit en est un exemple.</span><span class="sxs-lookup"><span data-stu-id="b8045-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="b8045-233">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="b8045-233">About this code, note:</span></span>

- <span data-ttu-id="b8045-234">Il surlignera uniquement les cellules qui disposent une valeur de nombre littérale.</span><span class="sxs-lookup"><span data-stu-id="b8045-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="b8045-235">Il surlignera les cellules qui disposent une formule (même si le résultat est un nombre) ou un booléen, un texte ou des cellules d’instruction d’erreur.</span><span class="sxs-lookup"><span data-stu-id="b8045-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="b8045-236">Pour tester le code, assurez-vous que la feuille de calcul dispose de certaines cellules avec des valeurs de nombre littérales, certaines avec d’autres sortes de valeurs littérales, et certaines avec des formules.</span><span class="sxs-lookup"><span data-stu-id="b8045-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="b8045-237">Parfois, vous avez besoin d’exécuter plus d’un type de valeur de cellule, tel que toutes les cellules à valeur de texte et à valeur booléen («Logique»).</span><span class="sxs-lookup"><span data-stu-id="b8045-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="b8045-238">L’enum`Excel.SpecialCellValueType` dispose de valeurs qui vous laisse combiner les types.</span><span class="sxs-lookup"><span data-stu-id="b8045-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="b8045-239">Par exemple, «LogicalText» ciblera toutes les cellules à valeur texte et booléen.</span><span class="sxs-lookup"><span data-stu-id="b8045-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="b8045-240">Vous pouvez combiner deux ou trois des quatre types de base.</span><span class="sxs-lookup"><span data-stu-id="b8045-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="b8045-241">Les noms de ces valeurs d’enum qui combinent les types de base sont toujours par ordre alphabétique.</span><span class="sxs-lookup"><span data-stu-id="b8045-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="b8045-242">Donc pour combiner les cellules à valeur d’erreur, texte et booléen, utilisez «ErrorLogicalText»,et non «LogicalErrorText» ou «TextErrorLogical».</span><span class="sxs-lookup"><span data-stu-id="b8045-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="b8045-243">Le paramètre par défaut de «Tous» combine les quatre types.</span><span class="sxs-lookup"><span data-stu-id="b8045-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="b8045-244">L’exemple suivant surligne toutes les cellules dotées de formules qui produisent les valeurs de nombre ou booléennes:</span><span class="sxs-lookup"><span data-stu-id="b8045-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="b8045-245">Le paramètre `Excel.SpecialCellValueType` peut uniquement être utilisé si le paramètre `Excel.SpecialCellType` est défini «Formules» ou «Constantes».</span><span class="sxs-lookup"><span data-stu-id="b8045-245">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="b8045-246">Obtenez RangeAreas dans RangeAreas</span><span class="sxs-lookup"><span data-stu-id="b8045-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="b8045-247">Le type `RangeAreas` lui-même dispose également de méthodes `getSpecialCells`et`getSpecialCellsOrNullObject` qui prennent les deux paramètres identiques.</span><span class="sxs-lookup"><span data-stu-id="b8045-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="b8045-248">Ces méthodes renvoient toutes les cellules ciblées à partir des plages dans la collection`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="b8045-249">Il existe une petite différence dans le comportement des méthodes lors de l’appel d’un objet`RangeAreas` au lieu d’un objet`Range`: lorsque vous passez «SameConditionalFormat» comme premier paramètre, la méthode renvoie toutes les cellules qui disposent la même mise en forme conditionnelle que la cellule en haut à gauche\* dans la première plage dans la `RangeAreas.areas`collection\*.</span><span class="sxs-lookup"><span data-stu-id="b8045-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="b8045-250">Le même point s’applique à «SameDataValidation»:lors du passage à`Range.getSpecialCells`, il renvoie toutes les cellules qui disposent la même règle de validation de données comme la cellule en haut à gauche*dans la plage*.</span><span class="sxs-lookup"><span data-stu-id="b8045-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="b8045-251">Mais lors du passage à `RangeAreas.getSpecialCells`, il renvoie toutes les cellules qui disposent la même règle de validation de données comme la cellule en haut à gauche*dans la plage`RangeAreas.areas`dans la collection*.</span><span class="sxs-lookup"><span data-stu-id="b8045-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="b8045-252">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="b8045-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="b8045-253">La lecture des valeurs de propriété de `RangeAreas` nécessite un soin, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes au sein du`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="b8045-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="b8045-254">La règle générales est que si une valeur consistante*peut*être renvoyée, elle sera renvoyée.</span><span class="sxs-lookup"><span data-stu-id="b8045-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="b8045-255">Par exemple, dans le code suivant, le code RGB pour rose (`#FFC0CB`) et`true`sera connecté à la console car les deux plages dans l’objet`RangeAreas` dispose d’un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="b8045-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="b8045-256">Les choses se compliquent lorsque la consistance est impossible.</span><span class="sxs-lookup"><span data-stu-id="b8045-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="b8045-257">Le comportement de propriétés`RangeAreas` suit ces trois principes:</span><span class="sxs-lookup"><span data-stu-id="b8045-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="b8045-258">Une propriété booléenne d’un objet`RangeAreas` renvoie`false`à moins que la propriété soit vraie pour toutes les plages membres.</span><span class="sxs-lookup"><span data-stu-id="b8045-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="b8045-259">Les propriétés non-booléennes, avec l’exception de la propriété`address`renvoie`null`à moins que la propriété correspondante sur toutes les plages membres dispose de la même valeur.</span><span class="sxs-lookup"><span data-stu-id="b8045-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="b8045-260">La propriété`address`renvoie une chaîne délimitée à virgule des adresses des plages membres.</span><span class="sxs-lookup"><span data-stu-id="b8045-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="b8045-261">Par exemple, le code suivante crée un`RangeAreas`dans lequel seule une plage est une colonne entière et seule une est remplie de rose.</span><span class="sxs-lookup"><span data-stu-id="b8045-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="b8045-262">La console s’affichera`null`pour un remplissage de couleur,`false`pour la propriété`isEntireRow` et «Sheet1!F3:F5, Sheet1!H:H» (en présumant que la feuille de calcule soit «Sheet1») pour la propriété`address`.</span><span class="sxs-lookup"><span data-stu-id="b8045-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="b8045-263">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b8045-263">See also</span></span>

- [<span data-ttu-id="b8045-264">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b8045-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="b8045-265">Objet de plage (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="b8045-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="b8045-266">[Objet RangeAreas (JavaScript API pout Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Ce lien peut ne peut pas fonctionner lorsque l’API est en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="b8045-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="b8045-267">Comme alternative, consultez [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="b8045-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>