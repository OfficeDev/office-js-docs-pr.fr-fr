---
title: Utiliser les classeurs utilisant l’API JavaScript Excel
description: ''
ms.date: 09/26/2019
localization_priority: Priority
ms.openlocfilehash: 66e531a382d467326e5132e60f06c98d414dbb16
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353873"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="90604-102">Utiliser les classeurs utilisant l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="90604-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="90604-103">Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de classeurs utilisant l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="90604-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="90604-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Classeur**, reportez-vous à la rubrique [Objet classeur (API JavaScript pour Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="90604-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="90604-105">Cet article décrit également les actions de niveau classeur effectuées via l’objet[Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="90604-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="90604-106">L’objet classeur est le point d’entrée pour votre complément pour interagir avec Excel.</span><span class="sxs-lookup"><span data-stu-id="90604-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="90604-107">Il gère les collections de feuilles de calcul, des tableaux, des tableaux croisés dynamiques et plus, via lesquels les données Excel sont consultées et modifiées.</span><span class="sxs-lookup"><span data-stu-id="90604-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="90604-108">L’objet[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) donne accès à votre complément aux données de tous les classeurs via les feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="90604-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="90604-109">Plus précisément, il permet à votre complément d’ajouter des feuilles de calcul et naviguer parmi celles-ci, et assigner des gestionnaires d’événements de feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="90604-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="90604-110">L’article [Manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md) décrit comment accéder et modifier des feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="90604-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="90604-111">Obtenir la cellule active ou la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="90604-111">Get the active cell or selected range</span></span>

<span data-ttu-id="90604-112">L’objet de classeur contient deux méthodes qui obtiennent une plage de cellules que l’utilisateur ou complément a sélectionnée : `getActiveCell()` et `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="90604-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="90604-113">`getActiveCell()` obtient la cellule active du classeur en tant qu’un [objet plage](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="90604-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="90604-114">L’exemple suivant montre un appel à `getActiveCell()`, suivi par adresse de la cellule imprimée sur la console.</span><span class="sxs-lookup"><span data-stu-id="90604-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="90604-115">Le `getSelectedRange()` méthode retourne la plage unique actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="90604-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="90604-116">Si plusieurs plages sont sélectionnées, une erreur InvalidSelection est envoyée.</span><span class="sxs-lookup"><span data-stu-id="90604-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="90604-117">L’exemple suivant montre un appel à `getSelectedRange()` qui définit ensuite la couleur de remplissage de la plage en jaune.</span><span class="sxs-lookup"><span data-stu-id="90604-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="90604-118">Créer un classeur</span><span class="sxs-lookup"><span data-stu-id="90604-118">Create a workbook</span></span>

<span data-ttu-id="90604-119">Votre complément peut créer un nouveau classeur, distinct de l’instance d’Excel dans laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="90604-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="90604-120">L’objet d’Excel a la méthode`createWorkbook` prévue à cet effet.</span><span class="sxs-lookup"><span data-stu-id="90604-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="90604-121">Lorsque cette méthode est appelée, le nouveau classeur est immédiatement ouvert et affiché dans une nouvelle instance d’Excel.</span><span class="sxs-lookup"><span data-stu-id="90604-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="90604-122">Votre complément reste ouvert et en cours d’exécution avec le classeur précédent.</span><span class="sxs-lookup"><span data-stu-id="90604-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="90604-123">La `createWorkbook` méthode peut également créer une copie d’un classeur existant.</span><span class="sxs-lookup"><span data-stu-id="90604-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="90604-124">La méthode accepte comme un paramètre facultatif une représentation de chaîne codée en base 64 d’un fichier .xlsx.</span><span class="sxs-lookup"><span data-stu-id="90604-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="90604-125">Le classeur résultant sera une copie de ce fichier, en supposant que l’argument de chaîne est un fichier .xlsx valide.</span><span class="sxs-lookup"><span data-stu-id="90604-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="90604-126">Vous pouvez accéder au classeur actif de votre complément en tant que chaîne codée en base 64 via [fichier découpage](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="90604-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="90604-127">La catégorie[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="90604-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="90604-128">Insérer une copie d’un classeur existant dans l’offre actuelle (préversion)</span><span class="sxs-lookup"><span data-stu-id="90604-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="90604-129">La méthode `WorksheetCollection.addFromBase64` est actuellement uniquement disponible en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="90604-129">The `WorksheetCollection.addFromBase64` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="90604-130">L’exemple précédent montre un nouveau classeur créé à partir d’un classeur existant.</span><span class="sxs-lookup"><span data-stu-id="90604-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="90604-131">Vous pouvez également copier la totalité ou une partie d’un classeur existant dans le tableau actuellement associé à votre complément.</span><span class="sxs-lookup"><span data-stu-id="90604-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="90604-132">Un classeur[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) a la `addFromBase64`méthode pour insérer des copies de feuilles de calcul du classeur cible dans lui-même.</span><span class="sxs-lookup"><span data-stu-id="90604-132">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="90604-133">Le fichier de l’autre classeur est passé en tant que chaîne codé en base 64, comme le `Excel.createWorkbook` appel.</span><span class="sxs-lookup"><span data-stu-id="90604-133">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="90604-134">L’exemple suivant montre des feuilles de calcul d’un classeur en cours d’insertion dans le classeur actif, juste après la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="90604-134">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="90604-135">Notez que`null` est passé pour le`sheetNamesToInsert?: string[]` paramètre.</span><span class="sxs-lookup"><span data-stu-id="90604-135">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="90604-136">Cela signifie que les feuilles de calcul sont insérées.</span><span class="sxs-lookup"><span data-stu-id="90604-136">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="90604-137">Protéger la structure du classeur</span><span class="sxs-lookup"><span data-stu-id="90604-137">Protect the workbook's structure</span></span>

<span data-ttu-id="90604-138">Votre complément permet de contrôler la possibilité d’un utilisateur de modifier la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="90604-138">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="90604-139">La propriété de l’objet classeur `protection` est un objet[WorkbookProtection](/javascript/api/excel/excel.workbookprotection) avec une méthode`protect()`.</span><span class="sxs-lookup"><span data-stu-id="90604-139">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="90604-140">L’exemple suivant illustre un scénario de base activer/désactiver la protection de la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="90604-140">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="90604-141">La méthode`protect` accepte un paramètre de chaîne facultatif.</span><span class="sxs-lookup"><span data-stu-id="90604-141">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="90604-142">Cette chaîne représente le mot de passe nécessaire pour un utilisateur pour ignorer la protection et modifier la structure du classeur.</span><span class="sxs-lookup"><span data-stu-id="90604-142">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="90604-143">La protection peut également être définie au niveau de la feuille de calcul pour empêcher la modification de données non souhaitée.</span><span class="sxs-lookup"><span data-stu-id="90604-143">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="90604-144">Pour plus d’informations, voir la section**protection des données** de l’article[manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="90604-144">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="90604-145">Pour plus d’informations sur la protection du classeur dans Excel, voir l’article [protéger un classeur](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="90604-145">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="90604-146">Accès aux propriétés du document</span><span class="sxs-lookup"><span data-stu-id="90604-146">Access document properties</span></span>

<span data-ttu-id="90604-147">Les objets classeur ont accès aux métadonnées de fichier Office, qui sont connues comme [propriétés du document](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="90604-147">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="90604-148">La propriété de l’objet classeur `properties` est un objet[DocumentProperties](/javascript/api/excel/excel.documentproperties) contenant ces valeurs de métadonnées.</span><span class="sxs-lookup"><span data-stu-id="90604-148">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="90604-149">L’exemple de code suivant montre comment définir la propriété d’**auteur**.
</span><span class="sxs-lookup"><span data-stu-id="90604-149">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="90604-150">Vous pouvez également définir des propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="90604-150">You can also define custom properties.</span></span> <span data-ttu-id="90604-151">L’objet DocumentProperties contient une propriété `custom` qui représente une collection de paires de valeur clés pour les propriétés définies par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="90604-151">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="90604-152">L’exemple suivant montre comment créer une propriété personnalisée nommée **Introduction** avec la valeur « Hello », puis la récupérer.</span><span class="sxs-lookup"><span data-stu-id="90604-152">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="90604-153">Accès aux paramètres de document</span><span class="sxs-lookup"><span data-stu-id="90604-153">Access document settings</span></span>

<span data-ttu-id="90604-154">Les paramètres d’un classeur sont similaires à la collection de propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="90604-154">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="90604-155">La différence est que les paramètres sont spécifiques à un seul fichier Excel et au jumelage complément, tandis que les propriétés sont uniquement connectées à celui-ci.</span><span class="sxs-lookup"><span data-stu-id="90604-155">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="90604-156">L’exemple suivant montre comment créer et accéder à un paramètre.</span><span class="sxs-lookup"><span data-stu-id="90604-156">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="90604-157">Ajouter des données XML personnalisées au classeur</span><span class="sxs-lookup"><span data-stu-id="90604-157">Add custom XML data to the workbook</span></span>

<span data-ttu-id="90604-158">Le format de fichier Open XML d’Excel **.xlsx** permet à votre complément d’incorporer des données XML personnalisées dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="90604-158">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="90604-159">Ces données continuent de s’afficher avec le classeur, indépendamment du complément.</span><span class="sxs-lookup"><span data-stu-id="90604-159">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="90604-160">Un classeur contient un[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), c'est-à-dire, une liste de[CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="90604-160">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="90604-161">Ceci octroie l’accès aux chaînes XML et ID correspondantes uniques.</span><span class="sxs-lookup"><span data-stu-id="90604-161">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="90604-162">En stockant ces ID en tant que paramètres, votre complément peut stocker les touches de ses parties XML entre les sessions.</span><span class="sxs-lookup"><span data-stu-id="90604-162">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="90604-163">Les exemples suivants montrent comment utiliser des éléments XML personnalisés.</span><span class="sxs-lookup"><span data-stu-id="90604-163">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="90604-164">Le premier bloc de code montre comment incorporer des données XML dans le document.</span><span class="sxs-lookup"><span data-stu-id="90604-164">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="90604-165">Il contient une liste de relecteurs, puis en utilisant les paramètres du classeur pour enregistrer le fichier XML`id` pour leur récupération future.</span><span class="sxs-lookup"><span data-stu-id="90604-165">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="90604-166">Le deuxième bloc montre comment accéder à ce XML ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="90604-166">The second block shows how to access that XML later.</span></span> <span data-ttu-id="90604-167">Le paramètre « ContosoReviewXmlPartId » est chargé et transmis au classeur`customXmlParts`.</span><span class="sxs-lookup"><span data-stu-id="90604-167">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="90604-168">Les données XML sont imprimées puis dans la console.</span><span class="sxs-lookup"><span data-stu-id="90604-168">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="90604-169">`CustomXMLPart.namespaceUri` est renseigné uniquement si l’élément XML personnalisé niveau supérieur contient l’attribut`xmlns`.</span><span class="sxs-lookup"><span data-stu-id="90604-169">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="90604-170">Contrôler le comportement de calcul</span><span class="sxs-lookup"><span data-stu-id="90604-170">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="90604-171">Définir le mode de calcul</span><span class="sxs-lookup"><span data-stu-id="90604-171">Set calculation mode</span></span>

<span data-ttu-id="90604-172">Par défaut, Excel recalcule les résultats d’une formule chaque fois qu’une cellule référencée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="90604-172">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="90604-173">Le performances de votre complément peuvent profiter de l’ajustement de ce comportement de calcul.</span><span class="sxs-lookup"><span data-stu-id="90604-173">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="90604-174">L’objet Application a une `calculationMode` propriété de type `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="90604-174">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="90604-175">Peut être défini à l'aide des valeurs suivantes :
</span><span class="sxs-lookup"><span data-stu-id="90604-175">It can be set to the following values:</span></span>

- <span data-ttu-id="90604-176">`automatic`: Le comportement de recalcul par défaut dans lequel Excel calcule les résultats d’une nouvelle formule chaque fois que les données pertinentes sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="90604-176">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="90604-177">`automaticExceptTables`: Identique `automatic`, sauf que les modifications apportées à des valeurs dans les tableaux sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="90604-177">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="90604-178">`manual`: Calculs sont uniquement effectués lorsque l’utilisateur ou un complément les demande.</span><span class="sxs-lookup"><span data-stu-id="90604-178">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="90604-179">Définir le type de calcul</span><span class="sxs-lookup"><span data-stu-id="90604-179">Set calculation type</span></span>

<span data-ttu-id="90604-180">L’objet [Application](/javascript/api/excel/excel.application) fournit une méthode pour forcer un nouveau calcul immédiat.</span><span class="sxs-lookup"><span data-stu-id="90604-180">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="90604-181">`Application.calculate(calculationType)` démarre un recalcul manuel basé sur la valeur `calculationType`.</span><span class="sxs-lookup"><span data-stu-id="90604-181">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="90604-182">Les valeurs suivantes peuvent être utilisées :</span><span class="sxs-lookup"><span data-stu-id="90604-182">The following values can be specified:</span></span>

- <span data-ttu-id="90604-183">`full`: Recalculer toutes les formules dans tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.</span><span class="sxs-lookup"><span data-stu-id="90604-183">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="90604-184">`fullRebuild`: Revérifier les formules dépendantes, puis recalculer toutes les formules de tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.</span><span class="sxs-lookup"><span data-stu-id="90604-184">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="90604-185">`recalculate`: Recalculer des formules qui ont changé (ou marqués par programme pour le recalcul) depuis le dernier calcul et les formules dépendantes, dans tous les classeurs actifs.</span><span class="sxs-lookup"><span data-stu-id="90604-185">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="90604-186">Pour plus d’informations sur le recalcul, voir l’article [recalcul de modification, l’itération ou la précision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="90604-186">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="90604-187">Suspendre temporairement les calculs</span><span class="sxs-lookup"><span data-stu-id="90604-187">Temporarily suspend calculations</span></span>

<span data-ttu-id="90604-188">L’API Excel vous permet également de désactiver les compléments calculs jusqu'à ce que `RequestContext.sync()` ne soit appelé.</span><span class="sxs-lookup"><span data-stu-id="90604-188">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="90604-189">Cette opération est effectuée avec `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="90604-189">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="90604-190">Utilisez cette méthode lorsque votre complément modifie de grandes plages sans avoir à accéder aux données entre les modifications.</span><span class="sxs-lookup"><span data-stu-id="90604-190">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="comments-preview"></a><span data-ttu-id="90604-191">Commentaires (préversion)</span><span class="sxs-lookup"><span data-stu-id="90604-191">Comments (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="90604-192">Les API de commentaire sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="90604-192">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="90604-193">Tous les [commentaires](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) d’un classeur sont suivis par la propriété `Workbook.comments`.</span><span class="sxs-lookup"><span data-stu-id="90604-193">All [comments](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="90604-194">Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément.</span><span class="sxs-lookup"><span data-stu-id="90604-194">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="90604-195">La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="90604-195">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span>

<span data-ttu-id="90604-196">Pour ajouter des commentaires à un classeur, utilisez la méthode `CommentCollection.add`, appliquez-la à la cellule dans laquelle le commentaire sera ajouté, sous forme de chaîne ou d’objet [Range](/javascript/api/excel/excel.range) et au texte du commentaire, sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="90604-196">To add comments to a workbook, use the `CommentCollection.add` method, passing in the cell where the comment will be added, as either a string or [Range](/javascript/api/excel/excel.range) object, and the comment's text, as a string.</span></span> <span data-ttu-id="90604-197">L’exemple de code suivant ajoute un commentaire à la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="90604-197">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("A2", "TODO: add data.");
    return context.sync();
});
```

<span data-ttu-id="90604-198">Chaque commentaire contient des métadonnées concernant sa création, notamment l’auteur et la date de création.</span><span class="sxs-lookup"><span data-stu-id="90604-198">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="90604-199">Les commentaires créés par votre complément sont considérés comme créés par l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="90604-199">Comments created by your add-in are considered to be authored by the current user.</span></span> <span data-ttu-id="90604-200">L’exemple suivant montre comment afficher l’adresse e-mail et le nom de l’auteur, ainsi que la date de création d’un commentaire dans la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="90604-200">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2.
    var comment = context.workbook.comments.getItemByCell("Comments!A2");
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

<span data-ttu-id="90604-201">Chaque commentaire contient zéro ou plusieurs réponses.</span><span class="sxs-lookup"><span data-stu-id="90604-201">Each comment contains zero or more replies.</span></span> <span data-ttu-id="90604-202">Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="90604-202">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="90604-203">Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse.</span><span class="sxs-lookup"><span data-stu-id="90604-203">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="90604-204">Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées.</span><span class="sxs-lookup"><span data-stu-id="90604-204">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="90604-205">L’exemple de code suivant ajoute une réponse au premier commentaire du classeur.</span><span class="sxs-lookup"><span data-stu-id="90604-205">The following code sample adds a data series to the first chart in the worksheet.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

<span data-ttu-id="90604-206">Pour modifier un commentaire ou une réponse à un commentaire, configurez sa propriété `Comment.content` ou `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="90604-206">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span> <span data-ttu-id="90604-207">Pour supprimer un commentaire ou une réponse à un commentaire, utilisez la méthode `Comment.delete` ou `CommentReply.delete`.</span><span class="sxs-lookup"><span data-stu-id="90604-207">To delete a comment or comment reply, use the `Comment.delete` method or `CommentReply.delete` method.</span></span> <span data-ttu-id="90604-208">La suppression d’un commentaire entraîne également celle de toutes les réponses associées à ce commentaire.</span><span class="sxs-lookup"><span data-stu-id="90604-208">Deleting a comment also deletes all the replies associated with that comment.</span></span>

> [!TIP]
> <span data-ttu-id="90604-209">Les commentaires peuvent également être gérés au niveau de la [feuille de calcul](/javascript/api/excel/excel.worksheet), en utilisant les mêmes techniques.</span><span class="sxs-lookup"><span data-stu-id="90604-209">Comments can also be managed at the [Worksheet](/javascript/api/excel/excel.worksheet) level using the same techniques.</span></span>

## <a name="save-the-workbook-preview"></a><span data-ttu-id="90604-210">Enregistrer le classeur (préversion)</span><span class="sxs-lookup"><span data-stu-id="90604-210">Save the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="90604-211">La méthode `Workbook.save` est actuellement uniquement disponible en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="90604-211">The `Workbook.save` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="90604-212">`Workbook.save` enregistre le classeur dans un espace de stockage permanent.</span><span class="sxs-lookup"><span data-stu-id="90604-212">`Workbook.save` saves the workbook to persistent storage .</span></span> <span data-ttu-id="90604-213">La méthode `save` accepte un paramètre `saveBehavior` unique et facultatif qui peut être l’une des valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="90604-213">The `save` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="90604-214">`Excel.SaveBehavior.save` (par défaut) : le fichier est enregistré sans inviter l’utilisateur à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="90604-214">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="90604-215">Si le fichier n’a pas été enregistré précédemment, il est enregistré dans l’emplacement par défaut.</span><span class="sxs-lookup"><span data-stu-id="90604-215">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="90604-216">Si le fichier a été enregistré précédemment, il est enregistré au même emplacement.</span><span class="sxs-lookup"><span data-stu-id="90604-216">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="90604-217">`Excel.SaveBehavior.prompt` : si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="90604-217">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="90604-218">Si le fichier a été enregistré précédemment, il est enregistré dans le même emplacement et l’utilisateur ne reçoit pas d’invite.</span><span class="sxs-lookup"><span data-stu-id="90604-218">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="90604-219">Si l’utilisateur est invité à enregistrer mais annule alors l’opération, `save` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="90604-219">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a><span data-ttu-id="90604-220">Fermer le classeur (préversion)</span><span class="sxs-lookup"><span data-stu-id="90604-220">Close the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="90604-221">La méthode `Workbook.close` est actuellement uniquement disponible en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="90604-221">The `Workbook.close` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="90604-222">`Workbook.close` ferme le classeur, ainsi que des compléments qui sont associées au classeur (l’application Excel reste ouverte).</span><span class="sxs-lookup"><span data-stu-id="90604-222">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="90604-223">La méthode `close` accepte un paramètre `closeBehavior` unique et facultatif qui peut être l’une des valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="90604-223">The `close` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="90604-224">`Excel.CloseBehavior.save` (par défaut) : le fichier est enregistré avant d’être fermé.</span><span class="sxs-lookup"><span data-stu-id="90604-224">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="90604-225">Si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="90604-225">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="90604-226">`Excel.CloseBehavior.skipSave` : le fichier est fermé immédiatement, sans enregistrer.</span><span class="sxs-lookup"><span data-stu-id="90604-226">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="90604-227">Les modifications non enregistrées sont perdues.</span><span class="sxs-lookup"><span data-stu-id="90604-227">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="90604-228">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="90604-228">See also</span></span>

- [<span data-ttu-id="90604-229">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="90604-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="90604-230">Utiliser les feuilles de calcul à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="90604-230">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="90604-231">Utiliser les plages à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="90604-231">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
