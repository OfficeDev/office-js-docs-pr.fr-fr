---
title: Travailler avec plusieurs plages simultanément dans les compléments Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016457"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="41dba-102">Travailler avec plusieurs plages simultanément dans les compléments Excel (Aperçu)</span><span class="sxs-lookup"><span data-stu-id="41dba-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="41dba-103">La bibliothèque JavaScript Excel permet à votre complément d'effectuer des opérations et de définir des propriétés sur plusieurs plages simultanément.</span><span class="sxs-lookup"><span data-stu-id="41dba-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="41dba-104">Les plages n’ont pas à être contiguës.</span><span class="sxs-lookup"><span data-stu-id="41dba-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="41dba-105">En plus de rendre votre code plus simple, cette méthode de définition de propriété s’exécute beaucoup plus rapidement que de définir la même propriété individuellement pour chacune des plages.</span><span class="sxs-lookup"><span data-stu-id="41dba-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="41dba-106">Les API décrites dans cet article nécessitent la **version Office 2016 Démarrer en un clic 1809 Build 10820.20000** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="41dba-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="41dba-107">(Vous devrez peut-être rejoindre le [programme Office Insider](https://products.office.com/office-insider) pour obtenir un build approprié). En outre, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir du [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="41dba-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="41dba-108">Enfin, nous n’avons pas encore les pages de référence de ces API.</span><span class="sxs-lookup"><span data-stu-id="41dba-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="41dba-109">Mais le fichier de type définition suivant comporte leurs descriptions : [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="41dba-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="41dba-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="41dba-110">RangeAreas</span></span>

<span data-ttu-id="41dba-111">Un ensemble de plages (éventuellement discontinu) est représenté par un objet `Excel.RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="41dba-112">Il possède des propriétés et méthodes similaires au type `Range` (beaucoup de noms identiques ou similaires), mais des ajustements ont été apportés :</span><span class="sxs-lookup"><span data-stu-id="41dba-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="41dba-113">Aux types de données pour les propriétés et le comportement des méthodes setter et getter.</span><span class="sxs-lookup"><span data-stu-id="41dba-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="41dba-114">Aux types de données des paramètres de méthode et des comportements de méthode.</span><span class="sxs-lookup"><span data-stu-id="41dba-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="41dba-115">Les types de données des valeurs renvoyées par les méthodes.</span><span class="sxs-lookup"><span data-stu-id="41dba-115">The data types of method return values.</span></span>

<span data-ttu-id="41dba-116">Voici quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="41dba-116">Some examples:</span></span>

- <span data-ttu-id="41dba-117">`RangeAreas` a une propriété `address` qui retourne une chaîne délimitée par des virgules d’adresses de plage, au lieu d’une seule adresse comme avec la propriété `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="41dba-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="41dba-118">`RangeAreas` a une propriété `dataValidation` qui retourne un objet `DataValidation` qui représente la validation des données de toutes les plages dans le `RangeAreas`, si elle est cohérente.</span><span class="sxs-lookup"><span data-stu-id="41dba-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="41dba-119">La propriété est `null` si des objets identiques `DataValidation` ne sont pas appliqués à toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="41dba-120">Il s’agit d’un principe général, mais pas universel, pour l'objet `RangeAreas` : *Si une propriété ne dispose pas de valeurs cohérentes sur toutes les plages dans le `RangeAreas`, elle est `null`.*</span><span class="sxs-lookup"><span data-stu-id="41dba-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="41dba-121">Pour plus d’informations et connaître certaines exceptions, voir [Propriétés de lecture de RangeAreas](#reading-properties-of-rangeareas).</span><span class="sxs-lookup"><span data-stu-id="41dba-121">See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="41dba-122">`RangeAreas.cellCount` obtient le nombre total de cellules dans toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="41dba-123">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="41dba-124">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` renvoient un autre objet `RangeAreas` qui représente toutes les colonnes (ou lignes) de toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="41dba-125">Par exemple, si le `RangeAreas` représente « A1: C4 » et « F14:L15 », puis `RangeAreas.getEntireColumn` renvoie un objet `RangeAreas` qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="41dba-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="41dba-126">`RangeAreas.copyFrom` peut adopter un paramètre `Range` ou un paramètre `RangeAreas` qui représente la ou les plages sources de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="41dba-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="41dba-127">Liste complète des membres de Range qui sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="41dba-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="41dba-128">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41dba-128">Properties</span></span>

<span data-ttu-id="41dba-129">Familiarisez-vous avec la [Lecture des propriétés de RangeAreas](#reading-properties-of-rangeareas) avant d’écrire le code qui lit les propriétés listées.</span><span class="sxs-lookup"><span data-stu-id="41dba-129">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="41dba-130">Il existe quelques subtilités quant aux valeurs renvoyées.</span><span class="sxs-lookup"><span data-stu-id="41dba-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="41dba-131">address</span><span class="sxs-lookup"><span data-stu-id="41dba-131">address</span></span>
- <span data-ttu-id="41dba-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="41dba-132">addressLocal</span></span>
- <span data-ttu-id="41dba-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="41dba-133">cellCount</span></span>
- <span data-ttu-id="41dba-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="41dba-134">conditionalFormats</span></span>
- <span data-ttu-id="41dba-135">context</span><span class="sxs-lookup"><span data-stu-id="41dba-135">context</span></span>
- <span data-ttu-id="41dba-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="41dba-136">dataValidation</span></span>
- <span data-ttu-id="41dba-137">format</span><span class="sxs-lookup"><span data-stu-id="41dba-137">format</span></span>
- <span data-ttu-id="41dba-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="41dba-138">isEntireColumn</span></span>
- <span data-ttu-id="41dba-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="41dba-139">isEntireRow</span></span>
- <span data-ttu-id="41dba-140">style</span><span class="sxs-lookup"><span data-stu-id="41dba-140">style</span></span>
- <span data-ttu-id="41dba-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="41dba-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="41dba-142">Méthodes</span><span class="sxs-lookup"><span data-stu-id="41dba-142">Methods</span></span>

<span data-ttu-id="41dba-143">Les méthodes de plage en préversion sont marquées comme telles.</span><span class="sxs-lookup"><span data-stu-id="41dba-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="41dba-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="41dba-144">calculate()</span></span>
- <span data-ttu-id="41dba-145">clear()</span><span class="sxs-lookup"><span data-stu-id="41dba-145">clear()</span></span>
- <span data-ttu-id="41dba-146">convertDataTypeToText() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="41dba-147">convertToLinkedDataType() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="41dba-148">copyFrom() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="41dba-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="41dba-149">getEntireColumn()</span></span>
- <span data-ttu-id="41dba-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="41dba-150">getEntireRow()</span></span>
- <span data-ttu-id="41dba-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="41dba-151">getIntersection()</span></span>
- <span data-ttu-id="41dba-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="41dba-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="41dba-153">getOffsetRange() (getOffsetRangeAreas nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="41dba-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="41dba-154">getSpecialCells() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="41dba-155">getSpecialCellsOrNullObject() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="41dba-156">getTables() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-156">getTables() (preview)</span></span>
- <span data-ttu-id="41dba-157">getUsedRange() (getUsedRangeAreas nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="41dba-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="41dba-158">getUsedRangeOrNullObject() (getUsedRangeAreasOrNullObject nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="41dba-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="41dba-159">load()</span><span class="sxs-lookup"><span data-stu-id="41dba-159">load()</span></span>
- <span data-ttu-id="41dba-160">set()</span><span class="sxs-lookup"><span data-stu-id="41dba-160">set\*</span></span>
- <span data-ttu-id="41dba-161">setDirty() (préversion)</span><span class="sxs-lookup"><span data-stu-id="41dba-161">setDirty() (preview)</span></span>
- <span data-ttu-id="41dba-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="41dba-162">toJSON()</span></span>
- <span data-ttu-id="41dba-163">track()</span><span class="sxs-lookup"><span data-stu-id="41dba-163">track</span></span>
- <span data-ttu-id="41dba-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="41dba-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="41dba-165">Méthodes et propriétés spécifiques à RangeArea</span><span class="sxs-lookup"><span data-stu-id="41dba-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="41dba-166">Le type `RangeAreas` contient certaines propriétés et méthodes qui ne sont pas comprises dans l'objet `Range`.</span><span class="sxs-lookup"><span data-stu-id="41dba-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="41dba-167">Vous trouverez ci-dessous une sélection ce celles-ci :</span><span class="sxs-lookup"><span data-stu-id="41dba-167">The following is a selection of them:</span></span>

- <span data-ttu-id="41dba-168">`areas`: Un objet `RangeCollection` qui contient toutes les plages représentées par l'objet `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="41dba-169">L'objet `RangeCollection` est également nouveau et est similaire à d’autres objets de la collection Excel.</span><span class="sxs-lookup"><span data-stu-id="41dba-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="41dba-170">Il possède une propriété `items` qui est un tableau des objets `Range` représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="41dba-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="41dba-171">`areaCount`: Nombre total de plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="41dba-172">`getOffsetRangeAreas`: Fonctionne exactement comme [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’un `RangeAreas` est retourné et qu’il contient des plages qui sont décalées à partir d’une des plages dans le `RangeAreas` d’origine.</span><span class="sxs-lookup"><span data-stu-id="41dba-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="41dba-173">Créer RangeAreas et définir les propriétés</span><span class="sxs-lookup"><span data-stu-id="41dba-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="41dba-174">Vous pouvez créer un objet `RangeAreas` de deux manières :</span><span class="sxs-lookup"><span data-stu-id="41dba-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="41dba-175">Appelez `Worksheet.getRanges()` et lui passer une chaîne avec des adresses de plage délimitées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="41dba-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="41dba-176">Si une plage que vous souhaitez inclure a été créée dans un [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="41dba-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="41dba-177">Appelez `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="41dba-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="41dba-178">Cette méthode renvoie un `RangeAreas` représentant toutes les plages qui sont sélectionnées dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="41dba-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="41dba-179">Une fois que vous avez un objet `RangeAreas`, vous pouvez en créer d’autres à l’aide des méthodes de l’objet qui renvoient `RangeAreas` telles que `getOffsetRangeAreas` et `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="41dba-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="41dba-180">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="41dba-181">Par exemple, la collection dans `RangeAreas.areas` ne possède pas de méthode `add`.</span><span class="sxs-lookup"><span data-stu-id="41dba-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="41dba-182">N’essayez pas de directement ajouter ou supprimer des membres du tableau `RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="41dba-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="41dba-183">Cela entraîne un comportement indésirable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="41dba-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="41dba-184">Par exemple, il est possible d'acheminer un objet `Range` supplémentaire sur le tableau, mais cela peut provoquer des erreurs, car les propriétés et méthodes `RangeAreas` se comportent comme si le nouvel élément n’est pas là.</span><span class="sxs-lookup"><span data-stu-id="41dba-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="41dba-185">Par exemple, la propriété `areaCount` n’inclut pas les plages poussées de cette manière et le `RangeAreas.getItemAt(index)` génère une erreur si `index` est supérieur(e) à `areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="41dba-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="41dba-186">De même, la suppression d'un objet `Range` dans le tableau `RangeAreas.areas.items` en en en obtenant une référence et en appelant sa méthode `Range.delete` provoque des bogues : bien que l'`Range`objet*est* supprimé, les propriétés et méthodes de l'objet parent `RangeAreas` se comportent ou essayent de se comporter, comme s’il était toujours présent.</span><span class="sxs-lookup"><span data-stu-id="41dba-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="41dba-187">Par exemple, si votre code appelle `RangeAreas.calculate`, Office tente de calculer la plage, mais une erreur se produit car l’objet range n’apparaît plus.</span><span class="sxs-lookup"><span data-stu-id="41dba-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="41dba-188">La définition d’une propriété sur une `RangeAreas` définit la propriété correspondante sur toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="41dba-189">Voici un exemple de définition d’une propriété sur plusieurs plages.</span><span class="sxs-lookup"><span data-stu-id="41dba-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="41dba-190">La fonction met en évidence les plages **F3:F5** et **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="41dba-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="41dba-191">Cet exemple s'applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de la plage que vous passez à `getRanges` ou facilement les calculer lors de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="41dba-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="41dba-192">Voici quelques scénarios où cela est possible :</span><span class="sxs-lookup"><span data-stu-id="41dba-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="41dba-193">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="41dba-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="41dba-194">Le code s’exécute dans le contexte de données importées pour lesquelles le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="41dba-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="41dba-195">Lorsque vous ne connaissez les plages sur lesquelles vous devez travailler au moment du codage, vous devez les découvrir lors de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="41dba-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="41dba-196">La section suivante décrit ces scénarios.</span><span class="sxs-lookup"><span data-stu-id="41dba-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="41dba-197">Découvrir les zones de plages par programme</span><span class="sxs-lookup"><span data-stu-id="41dba-197">Discover range areas programmatically</span></span>

<span data-ttu-id="41dba-198">Les méthodes `Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()` permettent de rechercher lors de l’exécution les plages que vous souhaitez utiliser en fonction des caractéristiques des cellules et du type des valeurs de cellules.</span><span class="sxs-lookup"><span data-stu-id="41dba-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="41dba-199">Voici les signatures des méthodes obtenues à partir du fichier de types de données TypeScript :</span><span class="sxs-lookup"><span data-stu-id="41dba-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="41dba-200">Voici un exemple d’utilisation de la première.</span><span class="sxs-lookup"><span data-stu-id="41dba-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="41dba-201">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="41dba-201">About this code, note:</span></span>

- <span data-ttu-id="41dba-202">Il limite la partie de la feuille qui doit être recherchée en appelant d’abord `Worksheet.getUsedRange`, puis en appelant `getSpecialCells` pour cette plage seulement.</span><span class="sxs-lookup"><span data-stu-id="41dba-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="41dba-203">Il passe en tant que paramètre à `getSpecialCells` la version de chaîne d’une valeur à partir de l'enum `Excel.SpecialCellType`.</span><span class="sxs-lookup"><span data-stu-id="41dba-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="41dba-204">Certaines valeurs qui peuvent être passées sont : « Blanks » pour les cellules vides, « Constants » pour les cellules contenant des valeurs littérales au lieu de formules et « SameConditionalFormat » pour les cellules qui ont la même mise en forme conditionnelle que la première cellule de la `usedRange`.</span><span class="sxs-lookup"><span data-stu-id="41dba-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="41dba-205">La première cellule est la cellule supérieure gauche.</span><span class="sxs-lookup"><span data-stu-id="41dba-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="41dba-206">Pour une liste complète des valeurs dans l'enum, voir [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="41dba-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="41dba-207">La méthode `getSpecialCells` renvoie un objet `RangeAreas`, de sorte que toutes les cellules contenant des formules seront colorés en rose, même si elles ne sont pas toutes contiguës.</span><span class="sxs-lookup"><span data-stu-id="41dba-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="41dba-208">Parfois la plage ne contient *aucune* cellule avec la caractéristique ciblée.</span><span class="sxs-lookup"><span data-stu-id="41dba-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="41dba-209">Si `getSpecialCells` ne trouve aucune cellule, elle génère une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="41dba-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="41dba-210">Cela redirige le flux de contrôle vers un bloc / une méthode `catch`, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="41dba-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="41dba-211">S'il n’en existe pas, l’erreur arrête la fonction.</span><span class="sxs-lookup"><span data-stu-id="41dba-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="41dba-212">Dans certains scénario, il est possible que vous vouliez justement qu'une erreur soit levée si aucune cellule avec la caractéristique ciblée n'existe.</span><span class="sxs-lookup"><span data-stu-id="41dba-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="41dba-213">Mais dans les scénarios dans lesquels cela est normal, même rare, qu’aucune cellule ne corresponde ; votre code doit prendre en compte cette possibilité et la traiter correctement sans lever une erreur.</span><span class="sxs-lookup"><span data-stu-id="41dba-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="41dba-214">Pour ces scénarios, utilisez la méthode `getSpecialCellsOrNullObject` et testez la propriété `RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="41dba-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="41dba-215">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="41dba-215">The following is an example.</span></span> <span data-ttu-id="41dba-216">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="41dba-216">Note about this code:</span></span>

- <span data-ttu-id="41dba-217">La méthode `getSpecialCellsOrNullObject` retourne toujours un objet proxy, elle n’est donc jamais `null` dans le sens JavaScript ordinaire.</span><span class="sxs-lookup"><span data-stu-id="41dba-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="41dba-218">Mais si aucune cellule n’est détectée, la propriété `isNullObject` de l’objet est définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="41dba-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="41dba-219">Elle appelle `context.sync` *avant* qu’il teste la propriété `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="41dba-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="41dba-220">Il s’agit d’une exigence de toutes les méthodes et propriétés `*OrNullObject`, car vous devez toujours charger et synchroniser une propriété afin de le lire.</span><span class="sxs-lookup"><span data-stu-id="41dba-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="41dba-221">Toutefois, il n’est pas nécessaire de charger *explicitement* la propriété `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="41dba-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="41dba-222">Elle est automatiquement chargée par le `context.sync` même si `load` n’est pas appelée sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="41dba-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="41dba-223">Pour plus d’informations, voir [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="41dba-223">For more information see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)</span></span>
- <span data-ttu-id="41dba-224">Vous pouvez tester ce code en sélection d'abord plage qui n’a aucune cellule de formule et en l’exécutant.</span><span class="sxs-lookup"><span data-stu-id="41dba-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="41dba-225">Sélectionnez ensuite une plage qui a au moins une cellule contenant une formule et exécutez-la à nouveau.</span><span class="sxs-lookup"><span data-stu-id="41dba-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

<span data-ttu-id="41dba-226">Par souci de simplicité, tous les autres exemples dans cet article utilisent la méthode `getSpecialCells` au lieu de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="41dba-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="41dba-227">Limiter les cellules cibles avec des types de valeur de cellule</span><span class="sxs-lookup"><span data-stu-id="41dba-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="41dba-228">Il existe un deuxième paramètre facultatif, de type enum `Excel.SpecialCellValueType`, qui permet de préciser encore le ciblage de cellules.</span><span class="sxs-lookup"><span data-stu-id="41dba-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="41dba-229">Vous pouvez l’utiliser uniquement lorsque vous passez « Formulas » ou « Constants » à `getSpecialCells` ou `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="41dba-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="41dba-230">Le paramètre spécifie que vous voulez uniquement les cellules avec certains types de valeurs.</span><span class="sxs-lookup"><span data-stu-id="41dba-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="41dba-231">Il existe quatre types de base : « Erreur », « Logique » (ce qui signifie booléenne), « Chiffres » et « Texte ».</span><span class="sxs-lookup"><span data-stu-id="41dba-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="41dba-232">(L’enum possède d'autres valeurs en plus de ces quatre qui sont présentées ci-dessous). Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="41dba-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="41dba-233">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="41dba-233">About this code, note:</span></span>

- <span data-ttu-id="41dba-234">Elle mettra uniquement en surbrillance les cellules qui ont une valeur numérique littérale.</span><span class="sxs-lookup"><span data-stu-id="41dba-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="41dba-235">Elle ne mettra pas en surbrillance les cellules qui contiennent une formule (même si le résultat est un nombre) ou une valeur booléenne, du texte ou les cellules d’état d’erreur.</span><span class="sxs-lookup"><span data-stu-id="41dba-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="41dba-236">Pour tester le code, assurez-vous que la feuille de calcul contienne des cellules avec des valeurs littérales numériques, d'autres avec d'autres types de valeurs littérales et certaines des formules.</span><span class="sxs-lookup"><span data-stu-id="41dba-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="41dba-237">Vous devez parfois travailler avec plus d’un type de valeur de cellule, telles qu'avec les cellules de texte et booléennes (« logical »).</span><span class="sxs-lookup"><span data-stu-id="41dba-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="41dba-238">L'enum `Excel.SpecialCellValueType` possède des valeurs qui vous permettent de combiner des types.</span><span class="sxs-lookup"><span data-stu-id="41dba-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="41dba-239">Par exemple, « LogicalText » ciblera toutes les cellules de type booléen et toutes les cellules de texte.</span><span class="sxs-lookup"><span data-stu-id="41dba-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="41dba-240">Vous pouvez combiner deux ou trois des quatre types de base.</span><span class="sxs-lookup"><span data-stu-id="41dba-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="41dba-241">Les noms de ces valeurs enum qui associent des types de base sont toujours dans l’ordre alphabétique.</span><span class="sxs-lookup"><span data-stu-id="41dba-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="41dba-242">Pour combiner des cellules d’erreur, de texte et des cellules booléennes, utilisez donc « ErrorLogicalText », pas « LogicalErrorText » ou « TextErrorLogical ».</span><span class="sxs-lookup"><span data-stu-id="41dba-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="41dba-243">Le paramètre par défaut de « All » regroupe les quatre types.</span><span class="sxs-lookup"><span data-stu-id="41dba-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="41dba-244">L’exemple suivant met en évidence toutes les cellules contenant des formules qui génèrent numéro ou valeurs booléennes :</span><span class="sxs-lookup"><span data-stu-id="41dba-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="41dba-245">Le paramètre `Excel.SpecialCellValueType` peut être utilisé uniquement si le paramètre `Excel.SpecialCellType` est « Formulas » ou « Constants ».</span><span class="sxs-lookup"><span data-stu-id="41dba-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="41dba-246">Obtenir RangeAreas dans RangeAreas</span><span class="sxs-lookup"><span data-stu-id="41dba-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="41dba-247">Le type `RangeAreas` lui-même contient également les méthodes `getSpecialCells` et `getSpecialCellsOrNullObject` qui comprennent les deux mêmes paramètres.</span><span class="sxs-lookup"><span data-stu-id="41dba-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="41dba-248">Ces méthodes retournent toutes les cellules ciblées à partir de toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="41dba-249">Il existe une petite différence dans le comportement des méthodes lorsqu’elles sont appelées sur un objet `RangeAreas` au lieu d’un objet `Range` : lorsque vous passez « SameConditionalFormat » en tant que le premier paramètre, la méthode renvoie toutes les cellules qui ont la même mise en forme conditionnelle que la cellule de gauche supérieure *dans la première plage de la `RangeAreas.areas` collection*.</span><span class="sxs-lookup"><span data-stu-id="41dba-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="41dba-250">Le même point s’applique à « SameDataValidation » : lorsqu’il est passé à `Range.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *de la plage*de cellules.</span><span class="sxs-lookup"><span data-stu-id="41dba-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="41dba-251">Mais lorsqu’il est passé à `RangeAreas.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *dans la première plage dans la `RangeAreas.areas` collection*.</span><span class="sxs-lookup"><span data-stu-id="41dba-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="41dba-252">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="41dba-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="41dba-253">Lire les valeurs de propriété de `RangeAreas` nécessite une attention particulière, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes dans la `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="41dba-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="41dba-254">La règle générale est que si une valeur cohérente *peut* être retournée, elle le sera.</span><span class="sxs-lookup"><span data-stu-id="41dba-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="41dba-255">Par exemple, dans le code suivant, le code RVB pour le rose (`#FFC0CB`) et `true` seront enregistrés dans la console, car à la fois les plages de l'objet `RangeAreas` ont un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="41dba-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

<span data-ttu-id="41dba-256">Lorsque la cohérence n’est pas possible les choses deviennent plus complexes.</span><span class="sxs-lookup"><span data-stu-id="41dba-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="41dba-257">Le comportement des propriétés `RangeAreas` suit ces trois principes :</span><span class="sxs-lookup"><span data-stu-id="41dba-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="41dba-258">Une propriété de type booléen d'un objet`RangeAreas` renvoie `false`, sauf si la propriété a la valeur true pour toutes les plages de membre.</span><span class="sxs-lookup"><span data-stu-id="41dba-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="41dba-259">Les propriétés non booléen, à l’exception de la propriété `address`, renvoient `null`, sauf si la propriété correspondante possède la même valeur sur toutes les plages de membre.</span><span class="sxs-lookup"><span data-stu-id="41dba-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="41dba-260">La propriété `address` renvoie une chaîne délimitée par des virgules des adresses des plages de membre.</span><span class="sxs-lookup"><span data-stu-id="41dba-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="41dba-261">Par exemple, le code suivant crée un `RangeAreas` dans lequel une seule plage est une colonne entière et une seule est rempli en rose.</span><span class="sxs-lookup"><span data-stu-id="41dba-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="41dba-262">La console affiche `null` pour la couleur de remplissage, `false` pour la propriété `isEntireRow` et « Sheet1!F3:F5, Sheet1!H:H» (en supposant que le nom de la feuille est « Sheet1 ») pour la propriété `address`.</span><span class="sxs-lookup"><span data-stu-id="41dba-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

## <a name="see-also"></a><span data-ttu-id="41dba-263">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="41dba-263">See also</span></span>

- [<span data-ttu-id="41dba-264">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="41dba-264">Excel JavaScript API core concepts</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="41dba-265">Objet Range (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="41dba-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="41dba-266">[Objet RangeAreas (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Il est possible que ce lien ne fonctionne pas si l’API est en préversion.</span><span class="sxs-lookup"><span data-stu-id="41dba-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="41dba-267">Comme alternative, voir [bêta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="41dba-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>