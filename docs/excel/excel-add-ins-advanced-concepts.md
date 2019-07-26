---
title: Concepts avancés de programmation avec l’API JavaScript Excel
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: 0270ca30e0add99dadc9fcfaf4a71cdb3fb68f46
ms.sourcegitcommit: 5e90a90175909e0f4f392f5c98bd1273f444fe49
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/24/2019
ms.locfileid: "35851587"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="92f9d-102">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="92f9d-103">Cet article s’appuie sur les informations contenues dans la rubrique [Concepts de base pour la programmation de l’API JavaScript Excel](excel-add-ins-core-concepts.md) pour décrire certains concepts plus avancés qui sont indispensables à la création de compléments complexes pour Excel 2016 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="92f9d-103">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="92f9d-104">API Office.js pour Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="92f9d-105">Un complément Excel interagit avec des objets dans Excel à l’aide de l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="92f9d-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="92f9d-106">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="92f9d-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="92f9d-107">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="92f9d-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="92f9d-108">Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune.</span><span class="sxs-lookup"><span data-stu-id="92f9d-108">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="92f9d-109">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="92f9d-109">For example:</span></span>

- <span data-ttu-id="92f9d-110">[Context](/javascript/api/office/office.context) : l’objet **Context** représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.</span><span class="sxs-lookup"><span data-stu-id="92f9d-110">[Context](/javascript/api/office/office.context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="92f9d-111">Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`.</span><span class="sxs-lookup"><span data-stu-id="92f9d-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="92f9d-112">En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="92f9d-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="92f9d-113">[Document](/javascript/api/office/office.document) : L’objet **Document** fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="92f9d-113">[Document](/javascript/api/office/office.document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

## <a name="requirement-sets"></a><span data-ttu-id="92f9d-114">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="92f9d-114">Requirement sets</span></span>

<span data-ttu-id="92f9d-p103">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour identifier les ensembles de conditions requises spécifiques disponibles sur chaque plateforme prise en charge, reportez-vous à [Ensembles de conditions requises de l’API JavaScript pour Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="92f9d-p103">Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="92f9d-118">Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="92f9d-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="92f9d-119">L’exemple de code suivant montre comment déterminer si l’application hôte dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.</span><span class="sxs-lookup"><span data-stu-id="92f9d-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3') === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="92f9d-120">Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="92f9d-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="92f9d-p104">Vous pouvez utiliser l’[élément Requirements](/office/dev/add-ins/reference/manifest/requirements) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément **Requirements** du manifeste, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans des compléments dans **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p104">You can use the [Requirements element](/office/dev/add-ins/reference/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="92f9d-123">L’exemple de code suivant montre l’élément **Requirements** dans un manifeste indiquant que le complément doit être chargé dans toutes les applications hôtes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="92f9d-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="92f9d-124">Pour rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel sur le web, Windows et iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="92f9d-124">To make your add-in available on all platforms of an Office host, such as Excel on Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="92f9d-125">Ensembles de conditions requises pour l’API commune Office.js</span><span class="sxs-lookup"><span data-stu-id="92f9d-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="92f9d-126">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="92f9d-126">For information about Common API requirement sets, see [Office Common API requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="92f9d-127">Chargement des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-127">Loading the properties of an object</span></span>

<span data-ttu-id="92f9d-p105">Tout appel de la méthode `load()` sur un objet JavaScript pour Excel demande à l’API de charger l’objet dans la mémoire JavaScript lorsque la méthode `sync()` est exécutée. La méthode `load()` accepte une chaîne qui contient des noms délimités par des virgules de propriétés à charger ou un objet spécifiant des propriétés à charger, des options de pagination, etc.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p105">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

> [!NOTE]
> <span data-ttu-id="92f9d-p106">si vous appelez la méthode `load()` sur un objet (ou collection) sans spécifier de paramètre, toutes les propriétés scalaires de l’objet (ou toutes les propriétés scalaires de tous les objets de la collection) sont chargées. Pour réduire la quantité de données transférées entre l’application hôte Excel et le complément, évitez d’appeler la méthode `load()` sans spécifier explicitement les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p106">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="92f9d-132">Détails de méthodes</span><span class="sxs-lookup"><span data-stu-id="92f9d-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="92f9d-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="92f9d-133">load(param: object)</span></span>

<span data-ttu-id="92f9d-134">Remplit l’objet de proxy créé dans le calque JavaScript avec les valeurs de propriété et d’objet spécifiées par les paramètres.</span><span class="sxs-lookup"><span data-stu-id="92f9d-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="92f9d-135">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="92f9d-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="92f9d-136">Paramètres</span><span class="sxs-lookup"><span data-stu-id="92f9d-136">Parameters</span></span>

|<span data-ttu-id="92f9d-137">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="92f9d-137">**Parameter**</span></span>|<span data-ttu-id="92f9d-138">**Type**</span><span class="sxs-lookup"><span data-stu-id="92f9d-138">**Type**</span></span>|<span data-ttu-id="92f9d-139">**Description**</span><span class="sxs-lookup"><span data-stu-id="92f9d-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="92f9d-140">objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-140">object</span></span>|<span data-ttu-id="92f9d-p107">Facultatif. Accepte des noms de propriétés sous forme de tableau ou de chaîne délimitée par des virgules. Un objet peut également être transmis pour définir les propriétés de sélection et de navigation (comme illustré dans l’exemple ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="92f9d-p107">Optional. Accepts parameter and relationship names as comma-delimited string or an array. An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="92f9d-144">Retourne</span><span class="sxs-lookup"><span data-stu-id="92f9d-144">Returns</span></span>

<span data-ttu-id="92f9d-145">void</span><span class="sxs-lookup"><span data-stu-id="92f9d-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="92f9d-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="92f9d-146">Example</span></span>

<span data-ttu-id="92f9d-p108">L’exemple de code suivant définit les propriétés d’une plage Excel en copiant les propriétés d’une autre plage. L’objet source doit d’abord être chargé, avant que ses valeurs de propriété puissent être accessibles et écrites sur la plage cible. L’exemple suppose que les deux plages (**B2:E2** et **B7:E7**) comprennent des données, et que leur mise en forme initiale est différente.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p108">The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="92f9d-150">Charger des propriétés d’option</span><span class="sxs-lookup"><span data-stu-id="92f9d-150">Load option properties</span></span>

<span data-ttu-id="92f9d-151">Au lieu de transmettre un tableau ou une chaîne délimitée par des virgules lorsque vous appelez la méthode `load()`, vous pouvez également transmettre un objet qui contient les propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="92f9d-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="92f9d-152">**Propriété**</span><span class="sxs-lookup"><span data-stu-id="92f9d-152">**Property**</span></span>|<span data-ttu-id="92f9d-153">**Type**</span><span class="sxs-lookup"><span data-stu-id="92f9d-153">**Type**</span></span>|<span data-ttu-id="92f9d-154">**Description**</span><span class="sxs-lookup"><span data-stu-id="92f9d-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="92f9d-155">objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-155">object</span></span>|<span data-ttu-id="92f9d-p109">Contient une liste délimitée par des virgules ou un tableau de propriétés scalaires. Facultatif.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="92f9d-158">objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-158">object</span></span>|<span data-ttu-id="92f9d-p110">Contient une liste délimitée par des virgules ou un tableau de propriétés de navigation. Facultatif.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="92f9d-161">int</span><span class="sxs-lookup"><span data-stu-id="92f9d-161">int</span></span>| <span data-ttu-id="92f9d-p111">Spécifie le nombre maximal d’éléments de collection qui peuvent être inclus dans le résultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="92f9d-165">int</span><span class="sxs-lookup"><span data-stu-id="92f9d-165">int</span></span>|<span data-ttu-id="92f9d-p112">Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du jeu de résultats démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="92f9d-170">L’exemple de code suivant charge une collection de feuilles de calcul en sélectionnant la propriété `name` et l’élément `address` de la plage utilisée pour chaque feuille de calcul dans la collection.</span><span class="sxs-lookup"><span data-stu-id="92f9d-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="92f9d-171">Il indique également que seules les cinq premières feuilles de calcul de la collection doivent être chargées.</span><span class="sxs-lookup"><span data-stu-id="92f9d-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="92f9d-172">Vous pouvez traiter l’ensemble suivant de cinq feuilles de calcul en spécifiant `top: 10` et `skip: 5` en tant que valeurs d’attribut.</span><span class="sxs-lookup"><span data-stu-id="92f9d-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="92f9d-173">Propriétés scalaires et de navigation</span><span class="sxs-lookup"><span data-stu-id="92f9d-173">Scalar and navigation properties</span></span>

<span data-ttu-id="92f9d-174">Il existe deux catégories de propriétés: **scalaire** et **de navigation**.</span><span class="sxs-lookup"><span data-stu-id="92f9d-174">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="92f9d-175">Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON.</span><span class="sxs-lookup"><span data-stu-id="92f9d-175">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="92f9d-176">Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont assignés, et non pas la propriété directement.</span><span class="sxs-lookup"><span data-stu-id="92f9d-176">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="92f9d-177">Par exemple, les membres `name` et `position` sur l’objet [Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que `protection` et `tables` sont des propriétés de navigation.</span><span class="sxs-lookup"><span data-stu-id="92f9d-177">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> <span data-ttu-id="92f9d-178">`prompt` sur l’objet [DataValidation] est un exemple de propriété scalaire qui doit être définie à l’aide d’un objet JSON (`dv.prompt = { title: "MyPrompt"}`), au lieu de définir les sous-propriétés (`dv.prompt.title = "MyPrompt" // will not set the title`).</span><span class="sxs-lookup"><span data-stu-id="92f9d-178">`prompt` on the [DataValidation] object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="92f9d-179">Propriétés scalaires et propriétés de navigation avec `object.load()`</span><span class="sxs-lookup"><span data-stu-id="92f9d-179">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="92f9d-180">Tout appel de la méthode `object.load()` sans paramètre spécifié charge toutes les propriétés scalaires de l’objet. Les propriétés de navigation de l’objet ne sont pas chargées.</span><span class="sxs-lookup"><span data-stu-id="92f9d-180">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="92f9d-181">En outre, les propriétés de navigation ne peuvent pas être chargées directement.</span><span class="sxs-lookup"><span data-stu-id="92f9d-181">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="92f9d-182">Au lieu de cela, vous devez utiliser la méthode `load()` pour référencer des propriétés scalaires individuelles au sein de la propriété de navigation de votre choix.</span><span class="sxs-lookup"><span data-stu-id="92f9d-182">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="92f9d-183">Par exemple, pour charger le nom de la police d’une plage, vous devez spécifier les propriétés de navigation **format** et **font** en tant que chemin d’accès à la propriété **name** :</span><span class="sxs-lookup"><span data-stu-id="92f9d-183">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="92f9d-p116">grâce à l’API JavaScript pour Excel, vous pouvez définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pouvez définir la taille de police pour une plage à l’aide de `someRange.format.font.size = 10;`. Il est inutile de charger la propriété avant de la configurer.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p116">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="92f9d-187">Définition des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-187">Setting properties of an object</span></span>

<span data-ttu-id="92f9d-p117">La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse. Au lieu de définir des propriétés individuelles à l’aide de chemins de navigation comme décrit ci-dessus, vous pouvez utiliser la méthode `object.set()` qui est disponible sur tous les objets de l’API JavaScript pour Excel. Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p117">Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="92f9d-p118">la méthode `set()` est implémentée uniquement pour les objets dans les API JavaScript pour Office propres à un hôte, telles que l’API JavaScript pour Excel. Les API communes (partagées) ne prennent pas en charge cette méthode.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p118">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="92f9d-193">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="92f9d-193">set (properties: object, options: object)</span></span>

<span data-ttu-id="92f9d-p119">Les propriétés de l’objet sur lequel la méthode est appelée sont définies sur les valeurs spécifiées par les propriétés correspondantes de l’objet transmis. Si le paramètre `properties` est un objet JavaScript, toute propriété de l’objet transmis qui correspond à une propriété en lecture seule dans l’objet sur lequel la méthode est appelée sera ignorée ou générera une exception, en fonction de la valeur du paramètre `options`.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="92f9d-196">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="92f9d-196">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="92f9d-197">Paramètres</span><span class="sxs-lookup"><span data-stu-id="92f9d-197">Parameters</span></span>

|<span data-ttu-id="92f9d-198">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="92f9d-198">**Parameter**</span></span>|<span data-ttu-id="92f9d-199">**Type**</span><span class="sxs-lookup"><span data-stu-id="92f9d-199">**Type**</span></span>|<span data-ttu-id="92f9d-200">**Description**</span><span class="sxs-lookup"><span data-stu-id="92f9d-200">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="92f9d-201">objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-201">object</span></span>|<span data-ttu-id="92f9d-202">Objet de même type Office.js que l’objet sur lequel la méthode est appelée ou objet JavaScript avec des noms et des types de propriétés reflétant la structure de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="92f9d-202">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="92f9d-203">objet</span><span class="sxs-lookup"><span data-stu-id="92f9d-203">object</span></span>|<span data-ttu-id="92f9d-p120">Facultatif. Peut être transmis uniquement si le premier paramètre est un objet JavaScript. L’objet peut contenir la propriété suivante : `throwOnReadOnly?: boolean` (La valeur par défaut est `true` : générer une erreur si l’objet JavaScript transmis inclut des propriétés en lecture seule.)</span><span class="sxs-lookup"><span data-stu-id="92f9d-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="92f9d-207">Retourne</span><span class="sxs-lookup"><span data-stu-id="92f9d-207">Returns</span></span>

<span data-ttu-id="92f9d-208">void</span><span class="sxs-lookup"><span data-stu-id="92f9d-208">void</span></span>

#### <a name="example"></a><span data-ttu-id="92f9d-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="92f9d-209">Example</span></span>

<span data-ttu-id="92f9d-p121">L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet **Range**. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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

## <a name="42ornullobject-methods"></a><span data-ttu-id="92f9d-212">Méthodes &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="92f9d-212">&#42;OrNullObject methods</span></span>

<span data-ttu-id="92f9d-p122">De nombreuses méthodes d’API JavaScript pour Excel renvoient une exception lorsque la condition de l’API n’est pas remplie. Par exemple, si vous tentez d’obtenir une feuille de calcul en spécifiant le nom d’une feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renvoie une exception `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p122">Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="92f9d-p123">Au lieu d’implémenter une logique complexe de gestion des exceptions pour des scénarios similaires, vous pouvez utiliser la variante de la méthode `*OrNullObject` disponible pour les différentes méthodes de l’API JavaScript pour Excel. Une méthode `*OrNullObject` renvoie un objet Null (pas l’élément JavaScript `null`) au lieu de lever une exception si l’élément spécifié n’existe pas. Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection telle que **Worksheets** pour tenter de récupérer un élément de la collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié, s’il existe. Sinon, elle renvoie un objet Null. L’objet Null renvoyé contient la propriété booléenne `isNullObject` que vous pouvez étudier pour déterminer l’existence de l’objet.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p123">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="92f9d-p124">L’exemple de code suivant tente de récupérer une feuille de calcul nommée « Data » à l’aide de la méthode `getItemOrNullObject()`. Si la méthode renvoie un objet Null, une nouvelle feuille doit être créée pour pouvoir réaliser des actions sur la feuille.</span><span class="sxs-lookup"><span data-stu-id="92f9d-p124">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="92f9d-222">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="92f9d-222">See also</span></span>

* [<span data-ttu-id="92f9d-223">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-223">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="92f9d-224">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-224">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="92f9d-225">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-225">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="92f9d-226">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="92f9d-226">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
