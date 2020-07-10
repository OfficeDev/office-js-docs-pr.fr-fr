---
title: Concepts avancés de programmation avec l’API JavaScript Excel
description: Découvrez comment un complément Excel interagit avec des objets dans Excel à l'aide des modèles d'objet API JavaScript pour Office.
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 81602f48231f20b50a454134bc789dfdee2bbc12
ms.sourcegitcommit: 4f2f1c0a8ee777a43bb28efa226684261f4c4b9f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081395"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="b5b56-103">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-103">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="b5b56-104">Cet article s’appuie sur les informations contenues dans la rubrique [Concepts de base pour la programmation de l’API JavaScript Excel](excel-add-ins-core-concepts.md) pour décrire certains concepts plus avancés qui sont indispensables à la création de compléments complexes pour Excel 2016 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b5b56-104">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="b5b56-105">API Office.js pour Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-105">Office.js APIs for Excel</span></span>

<span data-ttu-id="b5b56-106">Un complément Excel interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="b5b56-106">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="b5b56-107">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](../reference/overview/excel-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="b5b56-107">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="b5b56-108">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="b5b56-108">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="b5b56-109">Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune.</span><span class="sxs-lookup"><span data-stu-id="b5b56-109">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="b5b56-110">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b5b56-110">For example:</span></span>

- <span data-ttu-id="b5b56-111">[Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.</span><span class="sxs-lookup"><span data-stu-id="b5b56-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="b5b56-112">Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`.</span><span class="sxs-lookup"><span data-stu-id="b5b56-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="b5b56-113">En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="b5b56-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="b5b56-114">[Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="b5b56-114">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="b5b56-115">L’image suivante illustre les situations dans lesquelles vous pouvez utiliser l’API JavaScript Excel ou les API communes.</span><span class="sxs-lookup"><span data-stu-id="b5b56-115">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Image des différences entre l’API Excel et les API communes](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a><span data-ttu-id="b5b56-117">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b56-117">Requirement sets</span></span>

<span data-ttu-id="b5b56-118">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="b5b56-118">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="b5b56-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span><span class="sxs-lookup"><span data-stu-id="b5b56-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="b5b56-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b5b56-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="b5b56-121">Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="b5b56-121">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="b5b56-122">L’exemple de code suivant montre comment déterminer si l’application hôte dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.</span><span class="sxs-lookup"><span data-stu-id="b5b56-122">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="b5b56-123">Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="b5b56-123">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="b5b56-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span><span class="sxs-lookup"><span data-stu-id="b5b56-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="b5b56-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span><span class="sxs-lookup"><span data-stu-id="b5b56-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="b5b56-126">L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications hôtes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b5b56-126">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="b5b56-127">Pour rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel sur le web, Windows et iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="b5b56-127">To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="b5b56-128">Ensembles de conditions requises pour l’API commune Office.js</span><span class="sxs-lookup"><span data-stu-id="b5b56-128">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="b5b56-129">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b5b56-129">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="b5b56-130">Chargement des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-130">Loading the properties of an object</span></span>

<span data-ttu-id="b5b56-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span><span class="sxs-lookup"><span data-stu-id="b5b56-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="b5b56-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span><span class="sxs-lookup"><span data-stu-id="b5b56-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

### <a name="method-details"></a><span data-ttu-id="b5b56-133">Détails des méthodes</span><span class="sxs-lookup"><span data-stu-id="b5b56-133">Method details</span></span>

#### `load(propertyNames?: string | string[])`

<span data-ttu-id="b5b56-134">Files d’attente de la commande pour charger les propriétés de l’objet spécifié.</span><span class="sxs-lookup"><span data-stu-id="b5b56-134">Queues up a command to load the specified properties of the object.</span></span> <span data-ttu-id="b5b56-135">Vous devez contacter `context.sync()` avant de lire les propriétés.</span><span class="sxs-lookup"><span data-stu-id="b5b56-135">You must call `context.sync()` before reading the properties.</span></span>

#### <a name="syntax"></a><span data-ttu-id="b5b56-136">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b5b56-136">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="b5b56-137">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b5b56-137">Parameters</span></span>

|<span data-ttu-id="b5b56-138">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="b5b56-138">**Parameter**</span></span>|<span data-ttu-id="b5b56-139">**Type**</span><span class="sxs-lookup"><span data-stu-id="b5b56-139">**Type**</span></span>|<span data-ttu-id="b5b56-140">**Description**</span><span class="sxs-lookup"><span data-stu-id="b5b56-140">**Description**</span></span>|
|:------------|:-------|:----------|
|`propertyNames`|<span data-ttu-id="b5b56-141">objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-141">object</span></span>|<span data-ttu-id="b5b56-142">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="b5b56-142">Optional.</span></span> <span data-ttu-id="b5b56-143">Accepte les noms de propriétés sous forme d’un tableau ou de chaînes séparées par des virgules.</span><span class="sxs-lookup"><span data-stu-id="b5b56-143">Accepts property names as comma-delimited string or an array.</span></span>|

#### <a name="returns"></a><span data-ttu-id="b5b56-144">Retourne</span><span class="sxs-lookup"><span data-stu-id="b5b56-144">Returns</span></span>

<span data-ttu-id="b5b56-145">void</span><span class="sxs-lookup"><span data-stu-id="b5b56-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="b5b56-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="b5b56-146">Example</span></span>

<span data-ttu-id="b5b56-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span><span class="sxs-lookup"><span data-stu-id="b5b56-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="b5b56-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span><span class="sxs-lookup"><span data-stu-id="b5b56-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="b5b56-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span><span class="sxs-lookup"><span data-stu-id="b5b56-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a><span data-ttu-id="b5b56-150">Charger des propriétés d’option</span><span class="sxs-lookup"><span data-stu-id="b5b56-150">Load option properties</span></span>

<span data-ttu-id="b5b56-151">Au lieu de transmettre un tableau ou une chaîne délimitée par des virgules lorsque vous appelez la méthode `load()`, vous pouvez également transmettre un objet qui contient les propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b5b56-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="b5b56-152">**Propriété**</span><span class="sxs-lookup"><span data-stu-id="b5b56-152">**Property**</span></span>|<span data-ttu-id="b5b56-153">**Type**</span><span class="sxs-lookup"><span data-stu-id="b5b56-153">**Type**</span></span>|<span data-ttu-id="b5b56-154">**Description**</span><span class="sxs-lookup"><span data-stu-id="b5b56-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="b5b56-155">objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-155">object</span></span>|<span data-ttu-id="b5b56-156">Contains a comma-delimited list or an array of scalar property names.</span><span class="sxs-lookup"><span data-stu-id="b5b56-156">Contains a comma-delimited list or an array of scalar property names.</span></span> <span data-ttu-id="b5b56-157">Optional.</span><span class="sxs-lookup"><span data-stu-id="b5b56-157">Optional.</span></span>|
|`expand`|<span data-ttu-id="b5b56-158">objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-158">object</span></span>|<span data-ttu-id="b5b56-159">Contains a comma-delimited list or an array of navigational property names.</span><span class="sxs-lookup"><span data-stu-id="b5b56-159">Contains a comma-delimited list or an array of navigational property names.</span></span> <span data-ttu-id="b5b56-160">Optional.</span><span class="sxs-lookup"><span data-stu-id="b5b56-160">Optional.</span></span>|
|`top`|<span data-ttu-id="b5b56-161">int</span><span class="sxs-lookup"><span data-stu-id="b5b56-161">int</span></span>| <span data-ttu-id="b5b56-162">Specifies the maximum number of collection items that can be included in the result.</span><span class="sxs-lookup"><span data-stu-id="b5b56-162">Specifies the maximum number of collection items that can be included in the result.</span></span> <span data-ttu-id="b5b56-163">Optional.</span><span class="sxs-lookup"><span data-stu-id="b5b56-163">Optional.</span></span> <span data-ttu-id="b5b56-164">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="b5b56-164">You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="b5b56-165">int</span><span class="sxs-lookup"><span data-stu-id="b5b56-165">int</span></span>|<span data-ttu-id="b5b56-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span><span class="sxs-lookup"><span data-stu-id="b5b56-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span></span> <span data-ttu-id="b5b56-167">If `top` is specified, the result set will start after skipping the specified number of items.</span><span class="sxs-lookup"><span data-stu-id="b5b56-167">If `top` is specified, the result set will start after skipping the specified number of items.</span></span> <span data-ttu-id="b5b56-168">Optional.</span><span class="sxs-lookup"><span data-stu-id="b5b56-168">Optional.</span></span> <span data-ttu-id="b5b56-169">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="b5b56-169">You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="b5b56-170">L’exemple de code suivant charge une collection de feuilles de calcul en sélectionnant la propriété `name` et l’élément `address` de la plage utilisée pour chaque feuille de calcul dans la collection.</span><span class="sxs-lookup"><span data-stu-id="b5b56-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="b5b56-171">Il indique également que seules les cinq premières feuilles de calcul de la collection doivent être chargées.</span><span class="sxs-lookup"><span data-stu-id="b5b56-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="b5b56-172">Vous pouvez traiter l’ensemble suivant de cinq feuilles de calcul en spécifiant `top: 10` et `skip: 5` en tant que valeurs d’attribut.</span><span class="sxs-lookup"><span data-stu-id="b5b56-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

### <a name="calling-load-without-parameters"></a><span data-ttu-id="b5b56-173">Appel de `load` sans paramètres</span><span class="sxs-lookup"><span data-stu-id="b5b56-173">Calling `load` without parameters</span></span>

<span data-ttu-id="b5b56-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span><span class="sxs-lookup"><span data-stu-id="b5b56-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="b5b56-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span><span class="sxs-lookup"><span data-stu-id="b5b56-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b5b56-176">La quantité de données renvoyées par une `load`instruction sans paramètre peut dépasser les limites de taille du service.</span><span class="sxs-lookup"><span data-stu-id="b5b56-176">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="b5b56-177">Pour réduire les risques pesant sur les compléments plus anciens, certaines propriétés ne sont pas renvoyées par `load` sans en faire la demande explicite.</span><span class="sxs-lookup"><span data-stu-id="b5b56-177">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="b5b56-178">Les propriétés suivantes sont exclues des opérations de chargement suivantes :</span><span class="sxs-lookup"><span data-stu-id="b5b56-178">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="b5b56-179">Propriétés scalaires et de navigation</span><span class="sxs-lookup"><span data-stu-id="b5b56-179">Scalar and navigation properties</span></span>

<span data-ttu-id="b5b56-180">Il existe deux catégories de propriétés: **scalaire** et **de navigation**.</span><span class="sxs-lookup"><span data-stu-id="b5b56-180">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="b5b56-181">Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON.</span><span class="sxs-lookup"><span data-stu-id="b5b56-181">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="b5b56-182">Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont assignés, et non pas la propriété directement.</span><span class="sxs-lookup"><span data-stu-id="b5b56-182">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="b5b56-183">Par exemple, les membres `name` et `position` sur l’objet [Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que `protection` et `tables` sont des propriétés de navigation.</span><span class="sxs-lookup"><span data-stu-id="b5b56-183">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span> <span data-ttu-id="b5b56-184">`prompt` sur l’objet [DataValidation](/javascript/api/excel/excel.datavalidation) est un exemple de propriété scalaire qui doit être définie à l’aide d’un objet JSON (`dv.prompt = { title: "MyPrompt"}`), au lieu de définir les sous-propriétés (`dv.prompt.title = "MyPrompt" // will not set the title`).</span><span class="sxs-lookup"><span data-stu-id="b5b56-184">`prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="b5b56-185">Propriétés scalaires et propriétés de navigation avec `object.load()`</span><span class="sxs-lookup"><span data-stu-id="b5b56-185">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="b5b56-186">Tout appel de la méthode `object.load()` sans paramètre spécifié charge toutes les propriétés scalaires de l’objet. Les propriétés de navigation de l’objet ne sont pas chargées.</span><span class="sxs-lookup"><span data-stu-id="b5b56-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="b5b56-187">En outre, les propriétés de navigation ne peuvent pas être chargées directement.</span><span class="sxs-lookup"><span data-stu-id="b5b56-187">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="b5b56-188">Au lieu de cela, vous devez utiliser la méthode `load()` pour référencer des propriétés scalaires individuelles au sein de la propriété de navigation de votre choix.</span><span class="sxs-lookup"><span data-stu-id="b5b56-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="b5b56-189">Par exemple, pour charger le nom de la police d’une plage, vous devez spécifier les propriétés de navigation `format` et `font` en tant que chemin d’accès à la propriété `name` :</span><span class="sxs-lookup"><span data-stu-id="b5b56-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="b5b56-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span><span class="sxs-lookup"><span data-stu-id="b5b56-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="b5b56-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="b5b56-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="b5b56-192">You do not need to load the property before you set it.</span><span class="sxs-lookup"><span data-stu-id="b5b56-192">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="b5b56-193">Définition des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-193">Setting properties of an object</span></span>

<span data-ttu-id="b5b56-194">Setting properties on an object with nested navigation properties can be cumbersome.</span><span class="sxs-lookup"><span data-stu-id="b5b56-194">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="b5b56-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="b5b56-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="b5b56-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span><span class="sxs-lookup"><span data-stu-id="b5b56-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="b5b56-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="b5b56-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="b5b56-198">The common (shared) APIs do not support this method.</span><span class="sxs-lookup"><span data-stu-id="b5b56-198">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="b5b56-199">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="b5b56-199">set (properties: object, options: object)</span></span>

<span data-ttu-id="b5b56-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span><span class="sxs-lookup"><span data-stu-id="b5b56-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span></span> <span data-ttu-id="b5b56-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span><span class="sxs-lookup"><span data-stu-id="b5b56-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="b5b56-202">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b5b56-202">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="b5b56-203">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b5b56-203">Parameters</span></span>

|<span data-ttu-id="b5b56-204">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="b5b56-204">**Parameter**</span></span>|<span data-ttu-id="b5b56-205">**Type**</span><span class="sxs-lookup"><span data-stu-id="b5b56-205">**Type**</span></span>|<span data-ttu-id="b5b56-206">**Description**</span><span class="sxs-lookup"><span data-stu-id="b5b56-206">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="b5b56-207">objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-207">object</span></span>|<span data-ttu-id="b5b56-208">Objet de même type Office.js que l’objet sur lequel la méthode est appelée ou objet JavaScript avec des noms et des types de propriétés reflétant la structure de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="b5b56-208">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="b5b56-209">objet</span><span class="sxs-lookup"><span data-stu-id="b5b56-209">object</span></span>|<span data-ttu-id="b5b56-210">Optional.</span><span class="sxs-lookup"><span data-stu-id="b5b56-210">Optional.</span></span> <span data-ttu-id="b5b56-211">Can only be passed when the first parameter is a JavaScript object.</span><span class="sxs-lookup"><span data-stu-id="b5b56-211">Can only be passed when the first parameter is a JavaScript object.</span></span> <span data-ttu-id="b5b56-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span><span class="sxs-lookup"><span data-stu-id="b5b56-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="b5b56-213">Retourne</span><span class="sxs-lookup"><span data-stu-id="b5b56-213">Returns</span></span>

<span data-ttu-id="b5b56-214">void</span><span class="sxs-lookup"><span data-stu-id="b5b56-214">void</span></span>

#### <a name="example"></a><span data-ttu-id="b5b56-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="b5b56-215">Example</span></span>

<span data-ttu-id="b5b56-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span><span class="sxs-lookup"><span data-stu-id="b5b56-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span></span> <span data-ttu-id="b5b56-217">This example assumes that there is data in range **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="b5b56-217">This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a><span data-ttu-id="b5b56-218">Méthodes &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="b5b56-218">&#42;OrNullObject methods</span></span>

<span data-ttu-id="b5b56-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span><span class="sxs-lookup"><span data-stu-id="b5b56-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="b5b56-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span><span class="sxs-lookup"><span data-stu-id="b5b56-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="b5b56-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="b5b56-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="b5b56-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span><span class="sxs-lookup"><span data-stu-id="b5b56-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="b5b56-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span><span class="sxs-lookup"><span data-stu-id="b5b56-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="b5b56-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span><span class="sxs-lookup"><span data-stu-id="b5b56-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="b5b56-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span><span class="sxs-lookup"><span data-stu-id="b5b56-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="b5b56-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span><span class="sxs-lookup"><span data-stu-id="b5b56-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="b5b56-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span><span class="sxs-lookup"><span data-stu-id="b5b56-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a><span data-ttu-id="b5b56-228">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b5b56-228">See also</span></span>

* [<span data-ttu-id="b5b56-229">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="b5b56-230">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-230">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="b5b56-231">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-231">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="b5b56-232">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b5b56-232">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
