---
title: Concepts avancés de programmation avec l’API JavaScript Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 190eb65e45ce246009b6d85d378571bd2f451e0b
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459251"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="e96b4-102">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="e96b4-103">Cet article s’appuie sur les informations contenues dans la rubrique [Concepts de programmation fondamentaux de l’API JavaScript d’Excel](excel-add-ins-core-concepts.md) pour décrire certains des concepts les plus avancés qui sont essentiels à la création de compléments complexes pour Excel 2016 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="e96b4-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="e96b4-104">API Office.js pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="e96b4-105">Un complément Excel interagit avec des objets dans Excel à l’aide de l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="e96b4-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e96b4-106">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="e96b4-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="e96b4-107">**API communes** : incluses dans Office 2013, les API communes (également appelées [API partagées](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)) peuvent être utilisées pour accéder à des fonctionnalités, telles que l’interface utilisateur, les boîtes de dialogue et les paramètres du client, qui sont communes à plusieurs types d’applications hôtes, comme Word, Excel et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e96b4-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="e96b4-p101">Bien que vous utiliserez probablement l'API JavaScript d'Excel pour développer la majorité des fonctionnalités dans des compléments destinés à Excel 2016 ou une version ultérieure, vous utiliserez également des objets dans l’API partagée. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="e96b4-p101">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016, you'll also use objects in the Shared API. For example:</span></span>

- <span data-ttu-id="e96b4-p102">[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) : l’objet **Context** représente l’environnement d’exécution du complément et donne accès à des objets clés de l’API. Il se compose des détails de configuration de classeur, tels que `contentLanguage` et `officeTheme` et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode `requirements.isSetSupported()`, que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel où le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p102">[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="e96b4-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) : L’objet **Document** fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="e96b4-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="e96b4-114">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e96b4-114">Requirement sets</span></span>

<span data-ttu-id="e96b4-p103">Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Un complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si un hôte Office prend en charge les API dont le complément a besoin. Pour identifier les ensembles de conditions requises spécifiques qui sont disponibles sur chaque plate-forme prise en charge, voir [Ensembles de conditions requises de l’API JavaScript d’Excel](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="e96b4-p103">Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="e96b4-118">Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="e96b4-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="e96b4-119">L’exemple de code suivant montre comment déterminer si l’application hôte dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.</span><span class="sxs-lookup"><span data-stu-id="e96b4-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="e96b4-120">Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="e96b4-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="e96b4-p104">Vous pouvez utiliser l’[élément Requirements](https://docs.microsoft.com/javascript/office/manifest/requirements?view=office-js) dans le manifeste du complément pour spécifier les ensembles de conditions requises minimaux et/ou les méthodes API que votre complément doit activer. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les méthodes API qui sont spécifiés dans l’élément **Requirements** du manifeste, le complément ne sera pas exécuté dans cet hôte ou cette plateforme et ne s’affichera pas dans la liste des compléments répertoriés dans **Mes Compléments**.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p104">You can use the [Requirements element](https://docs.microsoft.com/javascript/office/manifest/requirements?view=office-js) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="e96b4-123">L’exemple de code suivant montre l’élément **Requirements** dans un manifeste indiquant que le complément doit être chargé dans toutes les applications hôtes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="e96b4-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="e96b4-124">pour rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel pour Windows, Excel Online et Excel pour iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="e96b4-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="e96b4-125">Ensembles de conditions requises pour l’API commune Office.js</span><span class="sxs-lookup"><span data-stu-id="e96b4-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="e96b4-126">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="e96b4-126">For information about common API requirement sets, see [Office common API requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="e96b4-127">Chargement des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-127">Loading the properties of an object</span></span>

<span data-ttu-id="e96b4-p105">Appeler la méthode `load()` sur un objet JavaScript Excel indique à l’API de charger l’objet en mémoire JavaScript lors de l’exécution de la méthode `sync()`. La méthode `load()` accepte une chaîne qui contient les noms de propriétés délimités par des virgules à charger ou un objet qui spécifie les propriétés à charger, les options de la pagination, etc.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p105">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="e96b4-p106">Si vous appelez la méthode `load()` sur un objet (ou une collection) sans spécifier aucun paramètre, toutes les propriétés scalaires de l’objet (ou toutes les propriétés scalaires de tous les objets de la collection) seront chargées. Pour réduire la quantité de transfert de données entre l’application hôte Excel et le complément, vous devez éviter d’appeler la méthode `load()` sans spécifier explicitement les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p106">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="e96b4-132">Détails de méthodes</span><span class="sxs-lookup"><span data-stu-id="e96b4-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="e96b4-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="e96b4-133">load(param: object)</span></span>

<span data-ttu-id="e96b4-134">Remplit l’objet proxy créé dans la couche JavaScript avec les valeurs de propriété et d’objet spécifiées par les paramètres.</span><span class="sxs-lookup"><span data-stu-id="e96b4-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="e96b4-135">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e96b4-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="e96b4-136">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e96b4-136">Parameters</span></span>

|<span data-ttu-id="e96b4-137">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="e96b4-137">**Parameter**</span></span>|<span data-ttu-id="e96b4-138">**Type**</span><span class="sxs-lookup"><span data-stu-id="e96b4-138">**Type**</span></span>|<span data-ttu-id="e96b4-139">**Description**</span><span class="sxs-lookup"><span data-stu-id="e96b4-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="e96b4-140">objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-140">object</span></span>|<span data-ttu-id="e96b4-p107">Facultatif. Accepte les noms de paramètre et de relation en tant que chaîne délimitée par des virgules ou que tableau. Un objet peut également être passé pour définir les propriétés de sélection et de navigation (comme illustré dans l’exemple ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="e96b4-p107">Optional. Accepts parameter and relationship names as comma-delimited string or an array. An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="e96b4-144">Renvoie</span><span class="sxs-lookup"><span data-stu-id="e96b4-144">Returns</span></span>

<span data-ttu-id="e96b4-145">annuler</span><span class="sxs-lookup"><span data-stu-id="e96b4-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="e96b4-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="e96b4-146">Example</span></span>

<span data-ttu-id="e96b4-p108">L’exemple de code suivant définit les propriétés d’une plage Excel en copiant les propriétés d’une autre plage. Notez que l’objet source doit être chargé en premier, avant que ses valeurs de propriété soient accessibles et écrites dans la plage cible. Cet exemple suppose qu’il existe des données dans deux plages (**B2:E2** et **B7:E7**) et que les deux plages sont initialement formatées différemment.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p108">The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="e96b4-150">Charger des propriétés d’option</span><span class="sxs-lookup"><span data-stu-id="e96b4-150">Load option properties</span></span>

<span data-ttu-id="e96b4-151">Au lieu de transmettre un tableau ou une chaîne délimitée par des virgules lorsque vous appelez la méthode `load()`, vous pouvez également transmettre un objet qui contient les propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="e96b4-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="e96b4-152">**Propriété**</span><span class="sxs-lookup"><span data-stu-id="e96b4-152">**Property**</span></span>|<span data-ttu-id="e96b4-153">**Type**</span><span class="sxs-lookup"><span data-stu-id="e96b4-153">**Type**</span></span>|<span data-ttu-id="e96b4-154">**Description**</span><span class="sxs-lookup"><span data-stu-id="e96b4-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="e96b4-155">objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-155">object</span></span>|<span data-ttu-id="e96b4-p109">Contient une liste délimitée par des virgules ou un tableau de noms de paramètres/relations. Facultatif.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="e96b4-158">objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-158">object</span></span>|<span data-ttu-id="e96b4-p110">Contient une liste délimitée par des virgules ou un tableau de noms de relations. Facultatif.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="e96b4-161">int</span><span class="sxs-lookup"><span data-stu-id="e96b4-161">int</span></span>| <span data-ttu-id="e96b4-p111">Spécifie le nombre maximal d’éléments de collection qui peuvent être inclus dans le résultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="e96b4-165">int</span><span class="sxs-lookup"><span data-stu-id="e96b4-165">int</span></span>|<span data-ttu-id="e96b4-p112">Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du jeu de résultats démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="e96b4-p113">L’exemple de code suivant charge une collection de feuilles de calcul en sélectionnant la propriété  `name` et le `address` de la plage utilisée pour chaque feuille de calcul de la collection. Il spécifie également que seules les cinq premières feuilles de calcul de la collection doivent être chargées. Vous pouvez traiter l’ensemble de cinq feuilles de calcul suivant en spécifiant `top: 10` et `skip: 5` en tant que valeurs d’attribut.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p113">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection. It also specifies that only the top five worksheets in the collection should be loaded. You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="e96b4-173">Propriétés scalaires et de navigation</span><span class="sxs-lookup"><span data-stu-id="e96b4-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="e96b4-p114">Dans la documentation de référence de l’API JavaScript d’Excel, vous pouvez remarquer que les membres d’objet sont regroupés en deux catégories : les **propriétés** et les **relations**. Une propriété d’un objet est un membre scalaire tel qu’une chaîne, un entier ou une valeur booléenne, alors qu’une relation de objet (également appelée propriété de navigation) est un membre qui est soit un objet ou une collection d’objets. Par exemple, les membres `name` et `position` de l’objet [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) sont des propriétés scalaires, tandis que `protection` et `tables` sont des relations (propriétés de navigation).</span><span class="sxs-lookup"><span data-stu-id="e96b4-p114">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**. A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects. For example, `name` and `position` members on the [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="e96b4-177">Propriétés scalaires et propriétés de navigation avec `object.load()`</span><span class="sxs-lookup"><span data-stu-id="e96b4-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="e96b4-p115">Appeler la méthode `object.load()` sans aucun paramètre spécifié charge toutes les propriétés scalaires de l’objet ; les propriétés de navigation de l’objet ne seront pas chargées. En outre, les propriétés de navigation ne peuvent pas être chargées directement. Au lieu de cela, vous devez utiliser la méthode `load()` pour référencer les propriétés scalaires individuelles au sein de la propriété de navigation de votre choix. Par exemple, pour charger le nom de la police pour une plage, vous devez spécifier les propriétés de navigation **format** et **font** en tant que chemin d’accès à la propriété **name** :</span><span class="sxs-lookup"><span data-stu-id="e96b4-p115">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded. Additionally, navigation properties cannot be loaded directly. Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property. For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="e96b4-p116">Avec l’API JavaScript d’Excel, vous pouvez définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pourriez définir la taille de police pour une plage à l’aide de `someRange.format.font.size = 10;`. Il est inutile de charger la propriété avant de la définir.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p116">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="e96b4-185">Définition des propriétés d’un objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-185">Setting properties of an object</span></span>

<span data-ttu-id="e96b4-p117">La définition des propriétés sur un objet avec des propriétés de navigation imbriquées peut être fastidieux. Au lieu de définir des propriétés individuelles à l’aide de chemins d’accès de navigation comme indiqué ci-dessus, vous pouvez utiliser la méthode `object.set()` disponible sur tous les objets dans l’interface API JavaScript d’Excel. Avec cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant un autre objet du même type Office.js ou un objet JavaScript avec des propriétés structurées comme les propriétés de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p117">Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="e96b4-p118">La méthode `set()` est implémentée uniquement pour les objets des APIs JavaScript Office spécifiques à l’hôte, telles que l’interface API JavaScript d’Excel. Les API communes (partagées) ne prennent pas charge cette méthode.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p118">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="e96b4-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="e96b4-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="e96b4-p119">Les propriétés de l’objet sur lequel la méthode est appelée sont définies sur les valeurs spécifiées par les propriétés correspondantes de l’objet transmis. Si le paramètre `properties` est un objet JavaScript, toute propriété de l’objet transmis qui correspond à une propriété en lecture seule dans l’objet sur lequel la méthode est appelée sera ignorée ou générera une exception, en fonction de la valeur du paramètre `options`.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="e96b4-194">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e96b4-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="e96b4-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e96b4-195">Parameters</span></span>

|<span data-ttu-id="e96b4-196">**Paramètre**</span><span class="sxs-lookup"><span data-stu-id="e96b4-196">**Parameter**</span></span>|<span data-ttu-id="e96b4-197">**Type**</span><span class="sxs-lookup"><span data-stu-id="e96b4-197">**Type**</span></span>|<span data-ttu-id="e96b4-198">**Description**</span><span class="sxs-lookup"><span data-stu-id="e96b4-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="e96b4-199">objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-199">object</span></span>|<span data-ttu-id="e96b4-200">Objet de même type Office.js que l’objet sur lequel la méthode est appelée ou objet JavaScript avec des noms et des types de propriétés reflétant la structure de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="e96b4-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="e96b4-201">objet</span><span class="sxs-lookup"><span data-stu-id="e96b4-201">object</span></span>|<span data-ttu-id="e96b4-p120">Facultatif. Peut être transmis uniquement si le premier paramètre est un objet JavaScript. L’objet peut contenir la propriété suivante : `throwOnReadOnly?: boolean` (La valeur par défaut est `true` : générer une erreur si l’objet JavaScript transmis inclut des propriétés en lecture seule.)</span><span class="sxs-lookup"><span data-stu-id="e96b4-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="e96b4-205">Renvoie</span><span class="sxs-lookup"><span data-stu-id="e96b4-205">Returns</span></span>

<span data-ttu-id="e96b4-206">annuler</span><span class="sxs-lookup"><span data-stu-id="e96b4-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="e96b4-207">Exemple</span><span class="sxs-lookup"><span data-stu-id="e96b4-207">Example</span></span>

<span data-ttu-id="e96b4-p121">L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet **Range**. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="e96b4-210">Méthodes \*OrNullObject</span><span class="sxs-lookup"><span data-stu-id="e96b4-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="e96b4-p122">De nombreuses méthodes de l’interface API JavaScript d’Excel renvoient une exception lorsque la condition de l’API n’est pas remplie. Par exemple, si vous tentez d’obtenir une feuille de calcul en spécifiant un nom de feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renverra une exception `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p122">Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="e96b4-p123">Au lieu d’implémenter des logiques de gestion d’exceptions complexes pour des scénarios comme celui-ci, vous pouvez utiliser la variante de méthode `*OrNullObject` qui est disponible pour plusieurs méthodes dans l’API Javascript d’Excel. Une méthode `*OrNullObject` retournera un objet null (pas le `null` Javascript) au lieu de générer une exception si l’élément spécifié n’existe pas. Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection comme **Worksheets** pour tenter de récupérer un élément dans une collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié s’il existe ; sinon, elle renvoie un objet null. L’objet null renvoyé contient la propriété booléenne `isNullObject` que vous pouvez évaluer pour déterminer si l’objet existe.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p123">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="e96b4-p124">L’exemple de code suivant tente de récupérer une feuille de calcul nommée « Data » à l’aide de la méthode `getItemOrNullObject()`. Si la méthode renvoie un objet null, une nouvelle feuille doit être créée avant que des actions puissent être menées sur la feuille.</span><span class="sxs-lookup"><span data-stu-id="e96b4-p124">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e96b4-220">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e96b4-220">See also</span></span>
 
* [<span data-ttu-id="e96b4-221">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-221">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="e96b4-222">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-222">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="e96b4-223">Optimisation des performances de l'API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-223">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="e96b4-224">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96b4-224">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
