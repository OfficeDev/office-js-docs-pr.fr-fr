---
title: Travailler avec plusieurs plages simultanément dans les compléments Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506293"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="19fd9-102">Travailler avec plusieurs plages simultanément dans les compléments Excel (Aperçu)</span><span class="sxs-lookup"><span data-stu-id="19fd9-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="19fd9-p101">La bibliothèque JavaScript d'Excel permet à votre complément d'effectuer des opérations et de définir les propriétés simultanément sur plusieurs plages. Les plages n’ont pas à être contiguës. En plus de rendre votre code plus simple, cette méthode de définition d’une propriété s’exécute beaucoup plus rapidement que de définir la même propriété individuellement pour chacune des plages.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p101">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="19fd9-p102">Les API décrites dans cet article nécessitent **Office 2016 Démarrer en un clic  version 1809 Build 10820.20000** ou une version ultérieure. (Vous devrez peut-être rejoindre le [programme Office Insider](https://products.office.com/office-insider) pour obtenir un build approprié). En outre, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir du [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Enfin, nous n’avons pas de pages de référence pour ces API pour le moment. Mais le fichier de type définition suivant comporte leurs descriptions :  [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="19fd9-p102">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="19fd9-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="19fd9-110">RangeAreas</span></span>

<span data-ttu-id="19fd9-p103">Un ensemble de plages (éventuellement discontinus) est représenté par un objet  `Excel.RangeAreas`. Il possède des propriétés et méthodes similaires au type  `Range` (beaucoup de noms identiques ou similaires), mais des ajustements ont été apportés :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p103">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="19fd9-113">Aux types de données pour les propriétés et le comportement des méthodes setter et getter.</span><span class="sxs-lookup"><span data-stu-id="19fd9-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="19fd9-114">Aux types de données des paramètres de méthode et des comportements de méthode.</span><span class="sxs-lookup"><span data-stu-id="19fd9-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="19fd9-115">Les types de données des valeurs renvoyées par les méthodes.</span><span class="sxs-lookup"><span data-stu-id="19fd9-115">The data types of method return values.</span></span>

<span data-ttu-id="19fd9-116">Voici quelques exemples :</span><span class="sxs-lookup"><span data-stu-id="19fd9-116">Some examples:</span></span>

- <span data-ttu-id="19fd9-117">`RangeAreas` a une propriété `address` qui retourne une chaîne délimitée par des virgules d’adresses de plage, au lieu d’une seule adresse comme avec la propriété `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="19fd9-p104">`RangeAreas` a une propriété `dataValidation` qui retourne un objet `DataValidation` qui représente la validation des données de toutes les plages dans le `RangeAreas`, si elles sont cohérentes. La propriété est `null` si des objets `DataValidation` identiques ne sont pas appliqués à toutes les plages dans le `RangeAreas`. Il s’agit d’un principe général, mais pas universel, pour l'objet `RangeAreas` : *Si une propriété ne dispose pas de valeurs cohérentes sur toutes les plages dans le `RangeAreas`, alors elle est `null`.* Pour plus d’informations et des exceptions, voir [Propriétés de lecture de RangeAreas](#reading-properties-of-rangeareas) .</span><span class="sxs-lookup"><span data-stu-id="19fd9-p104">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="19fd9-122">`RangeAreas.cellCount` obtient le nombre total de cellules dans toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="19fd9-123">`RangeAreas.calculate` recalcule les cellules de toutes les plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="19fd9-p105">`RangeAreas.getEntireColumn` et `RangeAreas.getEntireRow` renvoie un autre objet `RangeAreas` qui représente toutes les colonnes (ou lignes) de toutes les plages dans le `RangeAreas`. Par exemple, si le `RangeAreas` représente « A1 : C4 » et « F14:L15 », alors `RangeAreas.getEntireColumn` renvoie un objet `RangeAreas`qui représente « A:C » et « F:L ».</span><span class="sxs-lookup"><span data-stu-id="19fd9-p105">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="19fd9-126">`RangeAreas.copyFrom` peut adopter un paramètre `Range` ou un paramètre `RangeAreas` qui représente la ou les plages sources de l’opération de copie.</span><span class="sxs-lookup"><span data-stu-id="19fd9-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="19fd9-127">Liste complète des membres de Range qui sont également disponibles sur RangeAreas</span><span class="sxs-lookup"><span data-stu-id="19fd9-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="19fd9-128">Propriétés</span><span class="sxs-lookup"><span data-stu-id="19fd9-128">Properties</span></span>

<span data-ttu-id="19fd9-p106">Soyez familiarisé avec les [Propriétés de lecture de RangeAreas](#reading-properties-of-rangeareas) avant d’écrire le code qui lit les propriétés répertoriées. Il existe quelques subtilités quant aux valeurs renvoyées.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p106">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="19fd9-131">address</span><span class="sxs-lookup"><span data-stu-id="19fd9-131">address</span></span>
- <span data-ttu-id="19fd9-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="19fd9-132">addressLocal</span></span>
- <span data-ttu-id="19fd9-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="19fd9-133">cellCount</span></span>
- <span data-ttu-id="19fd9-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="19fd9-134">conditionalFormats</span></span>
- <span data-ttu-id="19fd9-135">context</span><span class="sxs-lookup"><span data-stu-id="19fd9-135">context</span></span>
- <span data-ttu-id="19fd9-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="19fd9-136">dataValidation</span></span>
- <span data-ttu-id="19fd9-137">format</span><span class="sxs-lookup"><span data-stu-id="19fd9-137">format</span></span>
- <span data-ttu-id="19fd9-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="19fd9-138">isEntireColumn</span></span>
- <span data-ttu-id="19fd9-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="19fd9-139">isEntireRow</span></span>
- <span data-ttu-id="19fd9-140">style</span><span class="sxs-lookup"><span data-stu-id="19fd9-140">style</span></span>
- <span data-ttu-id="19fd9-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="19fd9-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="19fd9-142">Méthodes</span><span class="sxs-lookup"><span data-stu-id="19fd9-142">Methods</span></span>

<span data-ttu-id="19fd9-143">Les méthodes de plage en préversion sont marquées comme telles.</span><span class="sxs-lookup"><span data-stu-id="19fd9-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="19fd9-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="19fd9-144">calculate()</span></span>
- <span data-ttu-id="19fd9-145">clear()</span><span class="sxs-lookup"><span data-stu-id="19fd9-145">clear()</span></span>
- <span data-ttu-id="19fd9-146">convertDataTypeToText() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="19fd9-147">convertToLinkedDataType() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="19fd9-148">copyFrom() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="19fd9-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="19fd9-149">getEntireColumn()</span></span>
- <span data-ttu-id="19fd9-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="19fd9-150">getEntireRow()</span></span>
- <span data-ttu-id="19fd9-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="19fd9-151">getIntersection()</span></span>
- <span data-ttu-id="19fd9-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="19fd9-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="19fd9-153">getOffsetRange() (getOffsetRangeAreas nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="19fd9-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="19fd9-154">getSpecialCells() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="19fd9-155">getSpecialCellsOrNullObject() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="19fd9-156">getTables() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-156">getTables() (preview)</span></span>
- <span data-ttu-id="19fd9-157">getUsedRange() (getUsedRangeAreas nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="19fd9-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="19fd9-158">getUsedRangeOrNullObject() (getUsedRangeAreasOrNullObject nommé sur l’objet RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="19fd9-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="19fd9-159">load()</span><span class="sxs-lookup"><span data-stu-id="19fd9-159">load()</span></span>
- <span data-ttu-id="19fd9-160">set()</span><span class="sxs-lookup"><span data-stu-id="19fd9-160">set\*</span></span>
- <span data-ttu-id="19fd9-161">setDirty() (préversion)</span><span class="sxs-lookup"><span data-stu-id="19fd9-161">setDirty() (preview)</span></span>
- <span data-ttu-id="19fd9-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="19fd9-162">toJSON()</span></span>
- <span data-ttu-id="19fd9-163">track()</span><span class="sxs-lookup"><span data-stu-id="19fd9-163">track</span></span>
- <span data-ttu-id="19fd9-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="19fd9-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="19fd9-165">Méthodes et propriétés spécifiques à RangeArea</span><span class="sxs-lookup"><span data-stu-id="19fd9-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="19fd9-p107">Le type `RangeAreas` comprend certaines propriétés et méthodes que l'objet `Range` n'a pas. Vous trouverez ci-dessous une sélection d'entre elles :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p107">The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them:</span></span>

- <span data-ttu-id="19fd9-p108">`areas`: Objet `RangeCollection` qui contient toutes les plages représentés par l'objet `RangeAreas`. L'objet `RangeCollection` est également nouveau et est similaire à d’autres objets de la collection Excel. Il a une propriété `items` qui est un tableau d'objets `Range` représentant les plages.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p108">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="19fd9-171">`areaCount`: Nombre total de plages dans le `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="19fd9-172">`getOffsetRangeAreas`: Fonctionne exactement comme [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), sauf qu’un `RangeAreas` est retourné et qu’il contient des plages qui sont décalées à partir d’une des plages dans le `RangeAreas` d’origine.</span><span class="sxs-lookup"><span data-stu-id="19fd9-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="19fd9-173">Créer RangeAreas et définir les propriétés</span><span class="sxs-lookup"><span data-stu-id="19fd9-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="19fd9-174">Vous pouvez créer un objet `RangeAreas` de deux manières simples :</span><span class="sxs-lookup"><span data-stu-id="19fd9-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="19fd9-p109">Appelez `Worksheet.getRanges()` et passez-lui une chaîne avec des adresses de plage délimitées par des virgules. Si une plage que vous souhaitez inclure a été créée dans un [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), vous pouvez inclure le nom, au lieu de l’adresse, dans la chaîne.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p109">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="19fd9-p110">Appelez `Workbook.getSelectedRanges()`. Cette méthode renvoie un `RangeAreas` représentant toutes les plages qui sont sélectionnées dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p110">Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="19fd9-179">Une fois que vous avez un objet `RangeAreas`, vous pouvez en créer d’autres à l’aide des méthodes de l’objet qui renvoient `RangeAreas` telles que `getOffsetRangeAreas` et `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="19fd9-p111">Vous ne pouvez pas ajouter directement des plages supplémentaires à un objet `RangeAreas`. Par exemple, la collection dans `RangeAreas.areas` n’a pas de méthode `add`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p111">You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="19fd9-p112">N’essayez pas d’ajouter ou de supprimer directement des membres du tableau `RangeAreas.areas.items`. Cela entraîne un comportement indésirable dans votre code. Par exemple, il est possible d'envoyer un objet `Range` supplémentaire sur le tableau, mais cela peut provoquer des erreurs, car les propriétés et méthodes  `RangeAreas` se comportent comme si le nouvel élément n’est pas là. Par exemple, la propriété `areaCount` n’inclut pas les plages envoyées de cette manière et le `RangeAreas.getItemAt(index)` génère une erreur si `index` est supérieure à `areasCount-1`. De même, la suppression d'un objet `Range` dans le tableau `RangeAreas.areas.items` en en obtenant une référence et en appelant sa méthode `Range.delete` provoque des bogues : bien que l'`Range`objet *est* supprimé, les propriétés et méthodes de l'objet parent `RangeAreas` se comportent, ou essayent de se comporter, comme s’il était toujours présent. Par exemple, si votre code appelle `RangeAreas.calculate`, Office tente de calculer la plage, mais une erreur se produira, car l’objet plage a été supprimé.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p112">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="19fd9-188">La définition d’une propriété sur une `RangeAreas` définit la propriété correspondante sur toutes les plages dans la collection `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="19fd9-p113">Voici un exemple de définition d’une propriété sur plusieurs plages. La fonction met en évidence les plages **F3:F5** et **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p113">The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="19fd9-p114">Cet exemple s'applique aux scénarios dans lesquels vous pouvez coder en dur les adresses de la plage que vous passez à `getRanges` ou facilement les calculer lors de l’exécution. Voici quelques scénarios où cela est possible :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p114">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="19fd9-193">Le code s’exécute dans le contexte d’un modèle connu.</span><span class="sxs-lookup"><span data-stu-id="19fd9-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="19fd9-194">Le code s’exécute dans le contexte de données importées pour lesquelles le schéma des données est connu.</span><span class="sxs-lookup"><span data-stu-id="19fd9-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="19fd9-p115">Lorsque vous ne connaissez les plages sur lesquelles vous devez travailler au moment du codage, vous devez les découvrir lors de l’exécution. La section suivante décrit ces scénarios.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p115">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="19fd9-197">Découvrir les zones de plages par programme</span><span class="sxs-lookup"><span data-stu-id="19fd9-197">Discover range areas programmatically</span></span>

<span data-ttu-id="19fd9-p116">Les méthodes `Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()` permettent de rechercher lors de l’exécution les plages que vous souhaitez utiliser en fonction des caractéristiques des cellules et du type des valeurs de cellules. Voici les signatures des méthodes issues du fichier de types de données TypeScript :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p116">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="19fd9-p117">Voici un exemple d’utilisation de la première. Concernant ce code, veuillez noter :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p117">The following is an example of using the first one. About this code, note:</span></span>

- <span data-ttu-id="19fd9-202">Il limite la partie de la feuille dans laquelle la recherche doit être effectuée en appelant d’abord `Worksheet.getUsedRange`, puis en appelant `getSpecialCells` pour cette plage seulement.</span><span class="sxs-lookup"><span data-stu-id="19fd9-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="19fd9-p118">Il passe en tant que paramètre à `getSpecialCells` la version de chaîne d’une valeur obtenue à partir de l'enum `Excel.SpecialCellType`. Parmi les autres valeurs qui peuvent être passés comptent la valeur « Blanks » pour les cellules vides, la valeur « Constants » pour les cellules contenant des valeurs littérales au lieu de formules et la valeur « SameConditionalFormat » pour les cellules qui ont la même mise en forme conditionnelle que la première cellule de `usedRange`. La première cellule est la cellule supérieure gauche. Pour obtenir une liste complète des valeurs de l’enum, voir [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="19fd9-p118">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="19fd9-207">La méthode `getSpecialCells` renvoie un objet `RangeAreas`, de sorte que toutes les cellules contenant des formules seront colorés en rose, même si elles ne sont pas toutes contiguës.</span><span class="sxs-lookup"><span data-stu-id="19fd9-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="19fd9-p119">Parfois la plage ne possède pas *du tout* de cellule avec la caractéristique ciblée. Si `getSpecialCells` n'en trouve pas, il génère une erreur **ItemNotFound** . Le flux de contrôle est alors dévié vers un bloc / une méthode `catch`, le cas échéant. S'il n’en existe pas, l’erreur arrête la fonction. Dans certains scénarios, il se peut que vous souhaitiez justement que l’erreur soit générée si aucune cellule avec la caractéristique ciblée n'est présente.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p119">Sometimes the range doesn't have *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="19fd9-p120">Mais dans les scénarios dans lesquels il est normal, mais peut-être rare, qu’aucune cellule correspondante n'existe, votre code doit vérifier cette possibilité et la traiter correctement sans générer d'erreur. Pour ces scénarios, vous devez utiliser la méthode `getSpecialCellsOrNullObject` et tester la propriété `RangeAreas.isNullObject`. Voici un exemple. Concernant ce code, veuillez noter :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p120">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property. The following is an example. Note about this code:</span></span>

- <span data-ttu-id="19fd9-p121">La méthode `getSpecialCellsOrNullObject` retourne toujours un objet proxy, elle n’est donc jamais `null` dans le sens ordinaire pour JavaScript. Mais si aucune cellule n’est détectée, la propriété `isNullObject` de l’objet est définie sur `true` .</span><span class="sxs-lookup"><span data-stu-id="19fd9-p121">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="19fd9-p122">Elle appelle `context.sync` *avant* que la propriété `isNullObject`  ne soit testée. Il s’agit d’une condition pour toutes les méthodes et propriétés `*OrNullObject`, car vous devez toujours charger et synchroniser une propriété afin de la lire. Toutefois, il n’est pas nécessaire de charger *explicitement* la propriété `isNullObject`. Elle est automatiquement chargée par `context.sync`, même si `load` n’est pas appelée sur l’objet. Pour plus d’informations, voir [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="19fd9-p122">It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="19fd9-p123">Vous pouvez tester ce code en sélectionnant d'abord une plage qui n’a aucune formule de cellules pour l'exécuter. Sélectionnez ensuite une plage qui a au moins une cellule contenant une formule et exécutez-le à nouveau.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p123">You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="19fd9-226">Par souci de simplicité, tous les autres exemples dans cet article utilisent la méthode `getSpecialCells` au lieu de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="19fd9-227">Limiter les cellules cibles avec des types de valeur de cellule</span><span class="sxs-lookup"><span data-stu-id="19fd9-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="19fd9-p124">Il existe un deuxième paramètre facultatif, de type enum  `Excel.SpecialCellValueType`, qui permet de préciser encore le ciblage de cellules. Vous pouvez l’utiliser uniquement lorsque vous passez « Formulas » ou « Constants » à  `getSpecialCells` ou `getSpecialCellsOrNullObject`. Le paramètre spécifie que vous voulez uniquement les cellules avec certains types de valeurs. Il existe quatre types de base : « Error », « Logical » (ce qui signifie booléenne), « Numbers » et « Text ». (L’enum possède d'autres valeurs en plus de ces quatre qui sont présentées ci-dessous). Voici un exemple. Concernant ce code, veuillez noter :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p124">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target. You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`. The parameter specifies that you only want cells with certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:</span></span>

- <span data-ttu-id="19fd9-p125">Il affiche uniquement en surbrillance les cellules qui ont une valeur numérique littérale. Il n'affiche pas en surbrillance les cellules qui contiennent une formule (même si le résultat est un nombre), une valeur booléenne, du texte ou une erreur.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p125">It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="19fd9-236">Pour tester le code, assurez-vous que la feuille de calcul contienne des cellules avec des valeurs littérales numériques, d'autres avec d'autres types de valeurs littérales et certaines des formules.</span><span class="sxs-lookup"><span data-stu-id="19fd9-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="19fd9-p126">Vous devez parfois travailler avec plus d’un type de valeur de cellule, telles qu'avec les cellules de texte et booléennes (« logical »). L'enum `Excel.SpecialCellValueType` possède des valeurs qui vous permettent de combiner des types. Vous pouvez combiner deux ou trois des quatre types de base. Les noms de ces valeurs enum qui associent des types de base sont toujours dans l’ordre alphabétique. Par exemple, « LogicalText » ciblera toutes les cellules de type booléen et toutes les cellules de texte. Pour combiner des cellules d’erreur, de texte et des cellules booléennes, utilisez donc « ErrorLogicalText », pas « LogicalErrorText » ou « TextErrorLogical ». L’exemple suivant met en évidence toutes les cellules contenant des formules qui génèrent numéro ou valeurs booléennes :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p126">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". The default parameter of "All" combines all four types. The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="19fd9-245">Le paramètre `Excel.SpecialCellValueType` peut être utilisé uniquement si le paramètre `Excel.SpecialCellType` est « Formulas » ou « Constants ».</span><span class="sxs-lookup"><span data-stu-id="19fd9-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="19fd9-246">Obtenir RangeAreas dans RangeAreas</span><span class="sxs-lookup"><span data-stu-id="19fd9-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="19fd9-p127">Le type  `RangeAreas` lui-même contient également les méthodes `getSpecialCells` and `getSpecialCellsOrNullObject`  qui comprennent les deux mêmes paramètres. Ces méthodes retournent toutes les cellules ciblées à partir de toutes les plages dans la collection `RangeAreas.areas`. Il existe une petite différence dans le comportement des méthodes lorsqu’elles sont appelées sur un objet `RangeAreas` au lieu d’un objet `Range`  : lorsque vous passez « SameConditionalFormat » en tant que premier paramètre, la méthode renvoie toutes les cellules qui ont la même mise en forme conditionnelle que la cellule supérieure gauche *dans la première plage de la collection `RangeAreas.areas`*. Le même point s’applique à « SameDataValidation » : lorsqu’elle est passée à `Range.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *de la plage* de cellules. Mais lorsqu’elle est passée à `RangeAreas.getSpecialCells`, elle renvoie toutes les cellules qui ont la même règle de validation de données que la cellule supérieure gauche *de la première plage dans la collection `RangeAreas.areas`*.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p127">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*. But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="19fd9-252">Lire les propriétés de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="19fd9-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="19fd9-p128">Lire les valeurs de propriété de `RangeAreas` nécessite une attention particulière, car une propriété donnée peut avoir des valeurs différentes pour des plages différentes dans la `RangeAreas`. La règle générale est que si une valeur cohérente *peut* être retournée, elle le sera. Par exemple, dans le code suivant, le code RVB pour le rose (`#FFC0CB`) et `true`  seront enregistrés dans la console, car à la fois les plages de l'objet `RangeAreas` ont un remplissage rose et les deux sont des colonnes entières.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p128">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="19fd9-p129">Lorsque la cohérence n’est pas possible, les choses deviennent plus complexes. Le comportement des propriétés `RangeAreas` suit ces trois principes :</span><span class="sxs-lookup"><span data-stu-id="19fd9-p129">Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="19fd9-258">Une propriété de type booléen d'un objet `RangeAreas` renvoie `false`, sauf si la propriété a la valeur true pour toutes les plages de membre.</span><span class="sxs-lookup"><span data-stu-id="19fd9-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="19fd9-259">Les propriétés non booléen, à l’exception de la propriété `address`, renvoient `null`, sauf si la propriété correspondante possède la même valeur sur toutes les plages de membre.</span><span class="sxs-lookup"><span data-stu-id="19fd9-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="19fd9-260">La propriété `address` renvoie une chaîne délimitée par des virgules des adresses des plages de membre.</span><span class="sxs-lookup"><span data-stu-id="19fd9-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="19fd9-p130">Par exemple, le code suivant crée un `RangeAreas` dans lequel une seule plage est une colonne entière et une seule est remplie de rose. La console affiche `null` pour la couleur de remplissage, `false` pour la propriété `isEntireRow`, et la feuille « Sheet1!F3:F5, Sheet1!H:H » (en supposant que le nom de la feuille est « Sheet1 ») pour la propriété  `address`.</span><span class="sxs-lookup"><span data-stu-id="19fd9-p130">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="19fd9-263">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="19fd9-263">See also</span></span>

- [<span data-ttu-id="19fd9-264">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="19fd9-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="19fd9-265">Objet Range (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="19fd9-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="19fd9-p131">[Objet RangeAreas (JavaScript API pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Ce lien ne fonctionne pas lorsque l’API est en mode Aperçu. Comme alternative, voir [office.d.ts bêta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="19fd9-p131">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview. As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>