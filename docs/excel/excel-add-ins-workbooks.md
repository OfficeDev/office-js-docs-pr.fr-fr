---
title: Utiliser les classeurs utilisant l’API JavaScript Excel
description: Découvrez comment effectuer des tâches courantes avec des workbooks ou des fonctionnalités au niveau de l’application à l’aide Excel API JavaScript.
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 48ceb882a7beea3fa3ca08216f3ee1dd82ba4fa9
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/09/2021
ms.locfileid: "52853982"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="f0670-103">Utiliser les classeurs utilisant l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="f0670-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="f0670-104">Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de classeurs utilisant l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="f0670-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="f0670-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Workbook` [Objet Workbook (interface API JavaScript pour Excel).](/javascript/api/excel/excel.workbook)</span><span class="sxs-lookup"><span data-stu-id="f0670-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="f0670-106">Cet article décrit également les actions de niveau classeur effectuées via l’objet[Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="f0670-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="f0670-107">L’objet classeur est le point d’entrée pour votre complément pour interagir avec Excel.</span><span class="sxs-lookup"><span data-stu-id="f0670-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="f0670-108">Il gère les collections de feuilles de calcul, des tableaux, des tableaux croisés dynamiques et plus, via lesquels les données Excel sont consultées et modifiées.</span><span class="sxs-lookup"><span data-stu-id="f0670-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="f0670-109">L’objet[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) donne accès à votre complément aux données de tous les classeurs via les feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="f0670-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="f0670-110">Plus précisément, il permet à votre complément d’ajouter des feuilles de calcul et naviguer parmi celles-ci, et assigner des gestionnaires d’événements de feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="f0670-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="f0670-111">L’article [Manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md) décrit comment accéder et modifier des feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="f0670-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="f0670-112">Obtenir la cellule active ou la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="f0670-112">Get the active cell or selected range</span></span>

<span data-ttu-id="f0670-113">L’objet de classeur contient deux méthodes qui obtiennent une plage de cellules que l’utilisateur ou complément a sélectionnée : `getActiveCell()` et `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="f0670-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="f0670-114">`getActiveCell()` obtient la cellule active du classeur en tant qu’un [objet plage](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="f0670-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="f0670-115">L’exemple suivant montre un appel à `getActiveCell()`, suivi par adresse de la cellule imprimée sur la console.</span><span class="sxs-lookup"><span data-stu-id="f0670-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f0670-116">Le `getSelectedRange()` méthode retourne la plage unique actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="f0670-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="f0670-117">Si plusieurs plages sont sélectionnées, une erreur InvalidSelection est envoyée.</span><span class="sxs-lookup"><span data-stu-id="f0670-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="f0670-118">L’exemple suivant montre un appel à `getSelectedRange()` qui définit ensuite la couleur de remplissage de la plage en jaune.</span><span class="sxs-lookup"><span data-stu-id="f0670-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="f0670-119">Créer un classeur</span><span class="sxs-lookup"><span data-stu-id="f0670-119">Create a workbook</span></span>

<span data-ttu-id="f0670-120">Votre complément peut créer un nouveau classeur, distinct de l’instance d’Excel dans laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f0670-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="f0670-121">L’objet d’Excel a la méthode`createWorkbook` prévue à cet effet.</span><span class="sxs-lookup"><span data-stu-id="f0670-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="f0670-122">Lorsque cette méthode est appelée, le nouveau classeur est immédiatement ouvert et affiché dans une nouvelle instance d’Excel.</span><span class="sxs-lookup"><span data-stu-id="f0670-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="f0670-123">Votre complément reste ouvert et en cours d’exécution avec le classeur précédent.</span><span class="sxs-lookup"><span data-stu-id="f0670-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="f0670-124">La `createWorkbook` méthode peut également créer une copie d’un classeur existant.</span><span class="sxs-lookup"><span data-stu-id="f0670-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="f0670-125">La méthode accepte comme un paramètre facultatif une représentation de chaîne codée en base 64 d’un fichier .xlsx.</span><span class="sxs-lookup"><span data-stu-id="f0670-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="f0670-126">Le classeur résultant sera une copie de ce fichier, en supposant que l’argument de chaîne est un fichier .xlsx valide.</span><span class="sxs-lookup"><span data-stu-id="f0670-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="f0670-127">Vous pouvez obtenir le classez actuel de votre add-in sous la forme d’une chaîne codée en base 64 à l’aide du [slicing de fichier.](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="f0670-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="f0670-128">La catégorie[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="f0670-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="f0670-129">Insérer une copie d’un classeur existant dans l’offre actuelle (préversion)</span><span class="sxs-lookup"><span data-stu-id="f0670-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f0670-130">La `Workbook.insertWorksheetsFromBase64` méthode est actuellement disponible uniquement en prévisualisation publique.</span><span class="sxs-lookup"><span data-stu-id="f0670-130">The `Workbook.insertWorksheetsFromBase64` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="f0670-131">L’exemple précédent montre un nouveau classeur créé à partir d’un classeur existant.</span><span class="sxs-lookup"><span data-stu-id="f0670-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="f0670-132">Vous pouvez également copier la totalité ou une partie d’un classeur existant dans le tableau actuellement associé à votre complément.</span><span class="sxs-lookup"><span data-stu-id="f0670-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="f0670-133">Un [workbook a](/javascript/api/excel/excel.workbook) la méthode pour insérer des copies des feuilles de calcul du `insertWorksheetsFromBase64` workbook cible dans lui-même.</span><span class="sxs-lookup"><span data-stu-id="f0670-133">A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="f0670-134">Le fichier de l’autre classeeur est transmis sous la forme d’une chaîne codée en base 64, comme `Excel.createWorkbook` l’appel.</span><span class="sxs-lookup"><span data-stu-id="f0670-134">The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call.</span></span> 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> <span data-ttu-id="f0670-135">La `insertWorksheetsFromBase64` méthode est prise en charge pour Excel sur Windows, Mac et le web.</span><span class="sxs-lookup"><span data-stu-id="f0670-135">The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web.</span></span> <span data-ttu-id="f0670-136">Il n’est pas pris en charge pour iOS.</span><span class="sxs-lookup"><span data-stu-id="f0670-136">It's not supported for iOS.</span></span> <span data-ttu-id="f0670-137">En outre, dans Excel sur le Web cette méthode ne prend pas en charge les feuilles de calcul source avec les éléments PivotTable, Chart, Comment ou Slicer.</span><span class="sxs-lookup"><span data-stu-id="f0670-137">Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements.</span></span> <span data-ttu-id="f0670-138">Si ces objets sont présents, la `insertWorksheetsFromBase64` méthode renvoie `UnsupportedFeature` l’erreur dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="f0670-138">If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web.</span></span> 

<span data-ttu-id="f0670-139">L’exemple de code suivant montre comment insérer des feuilles de calcul à partir d’un autre workbook dans le workbook actuel.</span><span class="sxs-lookup"><span data-stu-id="f0670-139">The following code sample shows how to insert worksheets from another workbook into the current workbook.</span></span> <span data-ttu-id="f0670-140">Cet exemple de code traite d’abord un fichier de classer avec un objet et extrait une chaîne codée en base 64, puis il insère cette chaîne codée en base 64 dans le classez [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) actuel.</span><span class="sxs-lookup"><span data-stu-id="f0670-140">This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook.</span></span> <span data-ttu-id="f0670-141">Les nouvelles feuilles de calcul sont insérées après la feuille de calcul nommée **Sheet1**.</span><span class="sxs-lookup"><span data-stu-id="f0670-141">The new worksheets are inserted after the worksheet named **Sheet1**.</span></span> <span data-ttu-id="f0670-142">Notez qu’il est transmis en tant que paramètre pour la `[]` [propriété InsertWorksheetOptions.sheetNamesToInsert.](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)</span><span class="sxs-lookup"><span data-stu-id="f0670-142">Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property.</span></span> <span data-ttu-id="f0670-143">Cela signifie que toutes les feuilles de calcul du manuel cible sont insérées dans le manuel en cours.</span><span class="sxs-lookup"><span data-stu-id="f0670-143">This means that all the worksheets from the target workbook are inserted into the current workbook.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="f0670-144">Protéger la structure du classeur</span><span class="sxs-lookup"><span data-stu-id="f0670-144">Protect the workbook's structure</span></span>

<span data-ttu-id="f0670-145">Votre complément permet de contrôler la possibilité d’un utilisateur de modifier la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="f0670-145">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="f0670-146">La propriété de l’objet classeur `protection` est un objet[WorkbookProtection](/javascript/api/excel/excel.workbookprotection) avec une méthode`protect()`.</span><span class="sxs-lookup"><span data-stu-id="f0670-146">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="f0670-147">L’exemple suivant illustre un scénario de base activer/désactiver la protection de la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="f0670-147">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f0670-148">La méthode`protect` accepte un paramètre de chaîne facultatif.</span><span class="sxs-lookup"><span data-stu-id="f0670-148">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="f0670-149">Cette chaîne représente le mot de passe nécessaire pour un utilisateur pour ignorer la protection et modifier la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="f0670-149">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="f0670-150">La protection peut également être définie au niveau de la feuille de calcul pour empêcher la modification de données non souhaitée.</span><span class="sxs-lookup"><span data-stu-id="f0670-150">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="f0670-151">Pour plus d’informations, voir la section **protection des données** de l’article [manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="f0670-151">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="f0670-152">Pour plus d’informations sur la protection du classeur dans Excel, voir l’article [protéger un classeur](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="f0670-152">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="f0670-153">Accès aux propriétés du document</span><span class="sxs-lookup"><span data-stu-id="f0670-153">Access document properties</span></span>

<span data-ttu-id="f0670-154">Les objets classeur ont accès aux métadonnées de fichier Office, qui sont connues comme [propriétés du document](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="f0670-154">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="f0670-155">La propriété de l’objet classeur `properties` est un objet[DocumentProperties](/javascript/api/excel/excel.documentproperties) contenant ces valeurs de métadonnées.</span><span class="sxs-lookup"><span data-stu-id="f0670-155">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="f0670-156">L’exemple suivant montre comment définir la `author` propriété.</span><span class="sxs-lookup"><span data-stu-id="f0670-156">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="f0670-157">Propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="f0670-157">Custom properties</span></span>

<span data-ttu-id="f0670-158">Vous pouvez également définir des propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f0670-158">You can also define custom properties.</span></span> <span data-ttu-id="f0670-159">L’objet DocumentProperties contient une propriété `custom` qui représente une collection de paires de valeur clés pour les propriétés définies par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f0670-159">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="f0670-160">L’exemple suivant montre comment créer une propriété personnalisée nommée **Introduction** avec la valeur « Hello », puis la récupérer.</span><span class="sxs-lookup"><span data-stu-id="f0670-160">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="f0670-161">Propriétés personnalisées au niveau de la feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="f0670-161">Worksheet-level custom properties</span></span>

<span data-ttu-id="f0670-162">Les propriétés personnalisées peuvent également être définies au niveau de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="f0670-162">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="f0670-163">Ces propriétés sont similaires aux propriétés personnalisées au niveau du document, sauf que la même clé peut être répétée dans différentes feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="f0670-163">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="f0670-164">L’exemple suivant montre comment créer une propriété personnalisée nommée **WorksheetGroup** avec la valeur « Alpha » dans la feuille de calcul actuelle, puis la récupérer.</span><span class="sxs-lookup"><span data-stu-id="f0670-164">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="f0670-165">Accès aux paramètres de document</span><span class="sxs-lookup"><span data-stu-id="f0670-165">Access document settings</span></span>

<span data-ttu-id="f0670-166">Les paramètres d’un classeur sont similaires à la collection de propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f0670-166">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="f0670-167">La différence est que les paramètres sont spécifiques à un seul fichier Excel et au jumelage complément, tandis que les propriétés sont uniquement connectées à celui-ci.</span><span class="sxs-lookup"><span data-stu-id="f0670-167">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="f0670-168">L’exemple suivant montre comment créer et accéder à un paramètre.</span><span class="sxs-lookup"><span data-stu-id="f0670-168">The following example shows how to create and access a setting.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="access-application-culture-settings"></a><span data-ttu-id="f0670-169">Accéder aux paramètres de culture d’application</span><span class="sxs-lookup"><span data-stu-id="f0670-169">Access application culture settings</span></span>

<span data-ttu-id="f0670-170">Un workbook a des paramètres de langue et de culture qui affectent l’affichage de certaines données.</span><span class="sxs-lookup"><span data-stu-id="f0670-170">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="f0670-171">Ces paramètres peuvent vous aider à trouver des données lorsque les utilisateurs de votre add-in partagent des workbooks dans différentes langues et cultures.</span><span class="sxs-lookup"><span data-stu-id="f0670-171">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="f0670-172">Votre add-in peut utiliser l’analyse de chaîne pour localiser le format des nombres, des dates et des heures en fonction des paramètres de culture système afin que chaque utilisateur voie les données dans le format de sa propre culture.</span><span class="sxs-lookup"><span data-stu-id="f0670-172">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="f0670-173">`Application.cultureInfo`définit les paramètres de culture système en tant [qu’objet CultureInfo.](/javascript/api/excel/excel.cultureinfo)</span><span class="sxs-lookup"><span data-stu-id="f0670-173">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="f0670-174">Il contient des paramètres tels que le séparateur décimal numérique ou le format de date.</span><span class="sxs-lookup"><span data-stu-id="f0670-174">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="f0670-175">Certains paramètres de culture peuvent être modifiés par le biais [Excel’interface utilisateur.](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)</span><span class="sxs-lookup"><span data-stu-id="f0670-175">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="f0670-176">Les paramètres système sont conservés dans `CultureInfo` l’objet.</span><span class="sxs-lookup"><span data-stu-id="f0670-176">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="f0670-177">Toutes les modifications locales sont conservées en tant [que propriétés](/javascript/api/excel/excel.application)au niveau de l’application, telles que `Application.decimalSeparator` .</span><span class="sxs-lookup"><span data-stu-id="f0670-177">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="f0670-178">L’exemple suivant modifie le caractère séparateur décimal d’une chaîne numérique de « , » au caractère utilisé par les paramètres système.</span><span class="sxs-lookup"><span data-stu-id="f0670-178">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="f0670-179">Ajouter des données XML personnalisées au classeur</span><span class="sxs-lookup"><span data-stu-id="f0670-179">Add custom XML data to the workbook</span></span>

<span data-ttu-id="f0670-180">Le format de fichier Open XML d’Excel **.xlsx** permet à votre complément d’incorporer des données XML personnalisées dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="f0670-180">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="f0670-181">Ces données continuent de s’afficher avec le classeur, indépendamment du complément.</span><span class="sxs-lookup"><span data-stu-id="f0670-181">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="f0670-182">Un classeur contient un[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), c'est-à-dire, une liste de[CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="f0670-182">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="f0670-183">Ceci octroie l’accès aux chaînes XML et ID correspondantes uniques.</span><span class="sxs-lookup"><span data-stu-id="f0670-183">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="f0670-184">En stockant ces ID en tant que paramètres, votre complément peut stocker les touches de ses parties XML entre les sessions.</span><span class="sxs-lookup"><span data-stu-id="f0670-184">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="f0670-185">Les exemples suivants montrent comment utiliser des éléments XML personnalisés.</span><span class="sxs-lookup"><span data-stu-id="f0670-185">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="f0670-186">Le premier bloc de code montre comment incorporer des données XML dans le document.</span><span class="sxs-lookup"><span data-stu-id="f0670-186">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="f0670-187">Il contient une liste de relecteurs, puis en utilisant les paramètres du classeur pour enregistrer le fichier XML`id` pour leur récupération future.</span><span class="sxs-lookup"><span data-stu-id="f0670-187">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="f0670-188">Le deuxième bloc montre comment accéder à ce XML ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="f0670-188">The second block shows how to access that XML later.</span></span> <span data-ttu-id="f0670-189">Le paramètre « ContosoReviewXmlPartId » est chargé et transmis au classeur`customXmlParts`.</span><span class="sxs-lookup"><span data-stu-id="f0670-189">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="f0670-190">Les données XML sont imprimées puis dans la console.</span><span class="sxs-lookup"><span data-stu-id="f0670-190">The XML data is then printed to the console.</span></span>

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="f0670-191">`CustomXMLPart.namespaceUri` est renseigné uniquement si l’élément XML personnalisé niveau supérieur contient l’attribut`xmlns`.</span><span class="sxs-lookup"><span data-stu-id="f0670-191">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="f0670-192">Contrôler le comportement de calcul</span><span class="sxs-lookup"><span data-stu-id="f0670-192">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="f0670-193">Définir le mode de calcul</span><span class="sxs-lookup"><span data-stu-id="f0670-193">Set calculation mode</span></span>

<span data-ttu-id="f0670-194">Par défaut, Excel recalcule les résultats d’une formule chaque fois qu’une cellule référencée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="f0670-194">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="f0670-195">Le performances de votre complément peuvent profiter de l’ajustement de ce comportement de calcul.</span><span class="sxs-lookup"><span data-stu-id="f0670-195">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="f0670-196">L’objet Application a une `calculationMode` propriété de type `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="f0670-196">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="f0670-197">Peut être défini à l'aide des valeurs suivantes :
</span><span class="sxs-lookup"><span data-stu-id="f0670-197">It can be set to the following values:</span></span>

- <span data-ttu-id="f0670-198">`automatic`: Le comportement de recalcul par défaut dans lequel Excel calcule les résultats d’une nouvelle formule chaque fois que les données pertinentes sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="f0670-198">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="f0670-199">`automaticExceptTables`: Identique `automatic`, sauf que les modifications apportées à des valeurs dans les tableaux sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="f0670-199">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="f0670-200">`manual`: Calculs sont uniquement effectués lorsque l’utilisateur ou un complément les demande.</span><span class="sxs-lookup"><span data-stu-id="f0670-200">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="f0670-201">Définir le type de calcul</span><span class="sxs-lookup"><span data-stu-id="f0670-201">Set calculation type</span></span>

<span data-ttu-id="f0670-202">L’objet [Application](/javascript/api/excel/excel.application) fournit une méthode pour forcer un nouveau calcul immédiat.</span><span class="sxs-lookup"><span data-stu-id="f0670-202">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="f0670-203">`Application.calculate(calculationType)` démarre un recalcul manuel basé sur la valeur `calculationType`.</span><span class="sxs-lookup"><span data-stu-id="f0670-203">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="f0670-204">Les valeurs suivantes peuvent être utilisées :</span><span class="sxs-lookup"><span data-stu-id="f0670-204">The following values can be specified:</span></span>

- <span data-ttu-id="f0670-205">`full`: Recalculer toutes les formules dans tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.</span><span class="sxs-lookup"><span data-stu-id="f0670-205">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="f0670-206">`fullRebuild`: Revérifier les formules dépendantes, puis recalculer toutes les formules de tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.</span><span class="sxs-lookup"><span data-stu-id="f0670-206">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="f0670-207">`recalculate`: Recalculer des formules qui ont changé (ou marqués par programme pour le recalcul) depuis le dernier calcul et les formules dépendantes, dans tous les classeurs actifs.</span><span class="sxs-lookup"><span data-stu-id="f0670-207">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="f0670-208">Pour plus d’informations sur le recalcul, voir l’article [recalcul de modification, l’itération ou la précision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="f0670-208">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="f0670-209">Suspendre temporairement les calculs</span><span class="sxs-lookup"><span data-stu-id="f0670-209">Temporarily suspend calculations</span></span>

<span data-ttu-id="f0670-210">L’API Excel vous permet également de désactiver les compléments calculs jusqu'à ce que `RequestContext.sync()` ne soit appelé.</span><span class="sxs-lookup"><span data-stu-id="f0670-210">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="f0670-211">Cette opération est effectuée avec `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="f0670-211">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="f0670-212">Utilisez cette méthode lorsque votre complément modifie de grandes plages sans avoir à accéder aux données entre les modifications.</span><span class="sxs-lookup"><span data-stu-id="f0670-212">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation-preview"></a><span data-ttu-id="f0670-213">Détecter l’activation d’un workbook (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f0670-213">Detect workbook activation (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f0670-214">`Workbook.onActivated`L’événement est actuellement disponible uniquement en prévisualisation publique.</span><span class="sxs-lookup"><span data-stu-id="f0670-214">The `Workbook.onActivated` event is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="f0670-215">Votre add-in peut détecter lorsqu’un workbook est activé.</span><span class="sxs-lookup"><span data-stu-id="f0670-215">Your add-in can detect when a workbook is activated.</span></span> <span data-ttu-id="f0670-216">Un workbook devient *inactif* lorsque l’utilisateur bascule le focus vers un autre workbook, vers une autre application ou (dans Excel sur le Web) vers un autre onglet du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="f0670-216">A workbook becomes *inactive* when the user switches focus to another workbook, to another application, or (in Excel on the web) to another tab of the web browser.</span></span> <span data-ttu-id="f0670-217">Un workbook est *activé lorsque* l’utilisateur renvoie le focus au workbook.</span><span class="sxs-lookup"><span data-stu-id="f0670-217">A workbook is *activated* when the user returns focus to the workbook.</span></span> <span data-ttu-id="f0670-218">L’activation du workbook peut déclencher des fonctions de rappel dans votre complément, telles que l’actualisation des données du workbook.</span><span class="sxs-lookup"><span data-stu-id="f0670-218">The workbook activation can trigger callback functions in your add-in, such as refreshing workbook data.</span></span>

<span data-ttu-id="f0670-219">Pour détecter lorsqu’un workbook est activé, inscrivez un [handler](excel-add-ins-events.md#register-an-event-handler) d’événements pour l’événement [onActivated](/javascript/api/excel/excel.workbook#onActivated) d’un workbook.</span><span class="sxs-lookup"><span data-stu-id="f0670-219">To detect when a workbook is activated, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#onActivated) event of a workbook.</span></span> <span data-ttu-id="f0670-220">Les handlers d’événements `onActivated` de l’événement reçoivent un [objet WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) lorsque l’événement se déclenche.</span><span class="sxs-lookup"><span data-stu-id="f0670-220">Event handlers for the `onActivated` event receive a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object when the event fires.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f0670-221">`onActivated`L’événement ne détecte pas lorsqu’un workbook est ouvert.</span><span class="sxs-lookup"><span data-stu-id="f0670-221">The `onActivated` event doesn't detect when a workbook is opened.</span></span> <span data-ttu-id="f0670-222">Cet événement détecte uniquement lorsqu’un utilisateur bascule le focus vers un workbook déjà ouvert.</span><span class="sxs-lookup"><span data-stu-id="f0670-222">This event only detects when a user switches focus back to an already open workbook.</span></span>

<span data-ttu-id="f0670-223">L’exemple de code suivant montre comment inscrire le handler d’événements et `onActivated` configurer une fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f0670-223">The following code sample shows how to register the `onActivated` event handler and set up a callback function.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## <a name="save-the-workbook"></a><span data-ttu-id="f0670-224">Enregistrer le classeur</span><span class="sxs-lookup"><span data-stu-id="f0670-224">Save the workbook</span></span>

<span data-ttu-id="f0670-225">`Workbook.save` enregistre le classeur dans un espace de stockage permanent.</span><span class="sxs-lookup"><span data-stu-id="f0670-225">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="f0670-226">La méthode `save` accepte un paramètre `saveBehavior` unique et facultatif qui peut être l’une des valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="f0670-226">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="f0670-227">`Excel.SaveBehavior.save` (par défaut) : le fichier est enregistré sans inviter l’utilisateur à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="f0670-227">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="f0670-228">Si le fichier n’a pas été enregistré précédemment, il est enregistré dans l’emplacement par défaut.</span><span class="sxs-lookup"><span data-stu-id="f0670-228">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="f0670-229">Si le fichier a été enregistré précédemment, il est enregistré au même emplacement.</span><span class="sxs-lookup"><span data-stu-id="f0670-229">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="f0670-230">`Excel.SaveBehavior.prompt` : si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="f0670-230">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="f0670-231">Si le fichier a été enregistré précédemment, il est enregistré dans le même emplacement et l’utilisateur ne reçoit pas d’invite.</span><span class="sxs-lookup"><span data-stu-id="f0670-231">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="f0670-232">Si l’utilisateur est invité à enregistrer mais annule alors l’opération, `save` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="f0670-232">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="f0670-233">Fermer le classeur</span><span class="sxs-lookup"><span data-stu-id="f0670-233">Close the workbook</span></span>

<span data-ttu-id="f0670-234">`Workbook.close` ferme le classeur, ainsi que des compléments qui sont associées au classeur (l’application Excel reste ouverte).</span><span class="sxs-lookup"><span data-stu-id="f0670-234">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="f0670-235">La méthode `close` accepte un paramètre `closeBehavior` unique et facultatif qui peut être l’une des valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="f0670-235">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="f0670-236">`Excel.CloseBehavior.save` (par défaut) : le fichier est enregistré avant d’être fermé.</span><span class="sxs-lookup"><span data-stu-id="f0670-236">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="f0670-237">Si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="f0670-237">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="f0670-238">`Excel.CloseBehavior.skipSave` : le fichier est fermé immédiatement, sans enregistrer.</span><span class="sxs-lookup"><span data-stu-id="f0670-238">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="f0670-239">Les modifications non enregistrées sont perdues.</span><span class="sxs-lookup"><span data-stu-id="f0670-239">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="f0670-240">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f0670-240">See also</span></span>

- [<span data-ttu-id="f0670-241">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="f0670-241">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f0670-242">Utiliser les feuilles de calcul à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="f0670-242">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
