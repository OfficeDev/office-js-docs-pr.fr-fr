---
title: Utiliser les classeurs utilisant l’API JavaScript Excel
description: Découvrez comment effectuer des tâches courantes avec des workbooks ou des fonctionnalités au niveau de l’application à l’aide Excel API JavaScript.
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ed63337aad322762019e8a51e3f1cc1c202db210
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938726"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Utiliser les classeurs utilisant l’API JavaScript Excel

Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de classeurs utilisant l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Workbook` [Objet Workbook (interface API JavaScript pour Excel).](/javascript/api/excel/excel.workbook) Cet article décrit également les actions de niveau classeur effectuées via l’objet[Application](/javascript/api/excel/excel.application).

L’objet classeur est le point d’entrée pour votre complément pour interagir avec Excel. Il gère les collections de feuilles de calcul, des tableaux, des tableaux croisés dynamiques et plus, via lesquels les données Excel sont consultées et modifiées. L’objet[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) donne accès à votre complément aux données de tous les classeurs via les feuilles de calcul individuelles. Plus précisément, il permet à votre complément d’ajouter des feuilles de calcul et naviguer parmi celles-ci, et assigner des gestionnaires d’événements de feuille de calcul. L’article [Manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md) décrit comment accéder et modifier des feuilles de calcul.

## <a name="get-the-active-cell-or-selected-range"></a>Obtenir la cellule active ou la plage sélectionnée

L’objet de classeur contient deux méthodes qui obtiennent une plage de cellules que l’utilisateur ou complément a sélectionnée : `getActiveCell()` et `getSelectedRange()`. `getActiveCell()` obtient la cellule active du classeur en tant qu’un [objet plage](/javascript/api/excel/excel.range). L’exemple suivant montre un appel à `getActiveCell()`, suivi par adresse de la cellule imprimée sur la console.

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

Le `getSelectedRange()` méthode retourne la plage unique actuellement sélectionnée. Si plusieurs plages sont sélectionnées, une erreur InvalidSelection est envoyée. L’exemple suivant montre un appel à `getSelectedRange()` qui définit ensuite la couleur de remplissage de la plage en jaune.

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a>Créer un classeur

Votre complément peut créer un nouveau classeur, distinct de l’instance d’Excel dans laquelle le complément est en cours d’exécution. L’objet d’Excel a la méthode`createWorkbook` prévue à cet effet. Lorsque cette méthode est appelée, le nouveau classeur est immédiatement ouvert et affiché dans une nouvelle instance d’Excel. Votre complément reste ouvert et en cours d’exécution avec le classeur précédent.

```js
Excel.createWorkbook();
```

La `createWorkbook` méthode peut également créer une copie d’un classeur existant. La méthode accepte comme un paramètre facultatif une représentation de chaîne codée en base 64 d’un fichier .xlsx. Le classeur résultant sera une copie de ce fichier, en supposant que l’argument de chaîne est un fichier .xlsx valide.

Vous pouvez obtenir le classez actuel de votre add-in sous la forme d’une chaîne codée en base 64 à l’aide du [slicing de fichier.](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) La classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a>Insérer une copie d’un classeur existant dans l’offre actuelle

L’exemple précédent montre un nouveau classeur créé à partir d’un classeur existant. Vous pouvez également copier la totalité ou une partie d’un classeur existant dans le tableau actuellement associé à votre complément. Un [workbook a](/javascript/api/excel/excel.workbook) la méthode pour insérer des copies des feuilles de calcul du `insertWorksheetsFromBase64` workbook cible dans lui-même. Le fichier de l’autre classeeur est transmis sous la forme d’une chaîne codée en base 64, tout comme `Excel.createWorkbook` l’appel. 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> La `insertWorksheetsFromBase64` méthode est prise en charge Excel sur Windows, Mac et le web. Il n’est pas pris en charge pour iOS. En outre, dans Excel sur le Web cette méthode ne prend pas en charge les feuilles de calcul source avec les éléments PivotTable, Chart, Comment ou Slicer. Si ces objets sont présents, la `insertWorksheetsFromBase64` méthode renvoie `UnsupportedFeature` l’erreur dans Excel sur le Web. 

L’exemple de code suivant montre comment insérer des feuilles de calcul à partir d’un autre workbook dans le workbook actuel. Cet exemple de code traite d’abord un fichier de classer avec un objet et extrait une chaîne codée en base 64, puis il insère cette chaîne codée en base 64 dans le classez en [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) cours. Les nouvelles feuilles de calcul sont insérées après la feuille de calcul nommée **Sheet1**. Notez qu’il est transmis en tant que paramètre pour la `[]` [propriété InsertWorksheetOptions.sheetNamesToInsert.](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) Cela signifie que toutes les feuilles de calcul du workbook cible sont insérées dans le manuel en cours.

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

## <a name="protect-the-workbooks-structure"></a>Protéger la structure du classeur

Votre complément permet de contrôler la possibilité d’un utilisateur de modifier la structure du classeur. La propriété de l’objet classeur `protection` est un objet[WorkbookProtection](/javascript/api/excel/excel.workbookprotection) avec une méthode`protect()`. L’exemple suivant illustre un scénario de base activer/désactiver la protection de la structure du classeur.

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

La méthode`protect` accepte un paramètre de chaîne facultatif. Cette chaîne représente le mot de passe nécessaire pour un utilisateur pour ignorer la protection et modifier la structure du classeur.

La protection peut également être définie au niveau de la feuille de calcul pour empêcher la modification de données non souhaitée. Pour plus d’informations, voir la section **protection des données** de l’article [manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md#data-protection).

> [!NOTE]
> Pour plus d’informations sur la protection du classeur dans Excel, voir l’article [protéger un classeur](https://support.microsoft.com/office/7e365a4d-3e89-4616-84ca-1931257c1517).

## <a name="access-document-properties"></a>Accès aux propriétés du document

Les objets classeur ont accès aux métadonnées de fichier Office, qui sont connues comme [propriétés du document](https://support.microsoft.com/office/21d604c2-481e-4379-8e54-1dd4622c6b75). La propriété de l’objet classeur `properties` est un objet[DocumentProperties](/javascript/api/excel/excel.documentproperties) contenant ces valeurs de métadonnées. L’exemple suivant montre comment définir la `author` propriété.

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a>Propriétés personnalisées

Vous pouvez également définir des propriétés personnalisées. L’objet DocumentProperties contient une propriété `custom` qui représente une collection de paires de valeur clés pour les propriétés définies par l’utilisateur. L’exemple suivant montre comment créer une propriété personnalisée nommée **Introduction** avec la valeur « Hello », puis la récupérer.

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

#### <a name="worksheet-level-custom-properties"></a>Propriétés personnalisées au niveau de la feuille de calcul

Les propriétés personnalisées peuvent également être définies au niveau de la feuille de calcul. Ces propriétés sont similaires aux propriétés personnalisées au niveau du document, sauf que la même clé peut être répétée dans différentes feuilles de calcul. L’exemple suivant montre comment créer une propriété personnalisée nommée **WorksheetGroup** avec la valeur « Alpha » dans la feuille de calcul actuelle, puis la récupérer.

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

## <a name="access-document-settings"></a>Accès aux paramètres de document

Les paramètres d’un classeur sont similaires à la collection de propriétés personnalisées. La différence est que les paramètres sont spécifiques à un seul fichier Excel et au jumelage complément, tandis que les propriétés sont uniquement connectées à celui-ci. L’exemple suivant montre comment créer et accéder à un paramètre.

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

## <a name="access-application-culture-settings"></a>Accéder aux paramètres de culture d’application

Un workbook a des paramètres de langue et de culture qui affectent l’affichage de certaines données. Ces paramètres peuvent vous aider à trouver des données lorsque les utilisateurs de votre add-in partagent des workbooks dans différentes langues et cultures. Votre add-in peut utiliser l’analyse de chaîne pour localiser le format des nombres, des dates et des heures en fonction des paramètres de culture système afin que chaque utilisateur voie les données dans le format de sa propre culture.

`Application.cultureInfo`définit les paramètres de culture système en tant [qu’objet CultureInfo.](/javascript/api/excel/excel.cultureinfo) Il contient des paramètres tels que le séparateur décimal numérique ou le format de date.

Certains paramètres de culture peuvent être modifiés par le [biais Excel’interface utilisateur.](https://support.microsoft.com/office/c093b545-71cb-4903-b205-aebb9837bd1e) Les paramètres système sont conservés dans `CultureInfo` l’objet. Toutes les modifications locales sont conservées en tant [que propriétés](/javascript/api/excel/excel.application)au niveau de l’application, telles que `Application.decimalSeparator` .

L’exemple suivant modifie le caractère séparateur décimal d’une chaîne numérique de « , » au caractère utilisé par les paramètres système.

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

## <a name="add-custom-xml-data-to-the-workbook"></a>Ajouter des données XML personnalisées au classeur

Le format de fichier Open XML d’Excel **.xlsx** permet à votre complément d’incorporer des données XML personnalisées dans le classeur. Ces données continuent de s’afficher avec le classeur, indépendamment du complément.

Un classeur contient un[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), c'est-à-dire, une liste de[CustomXmlParts](/javascript/api/excel/excel.customxmlpart). Ceci octroie l’accès aux chaînes XML et ID correspondantes uniques. En stockant ces ID en tant que paramètres, votre complément peut stocker les touches de ses parties XML entre les sessions.

Les exemples suivants montrent comment utiliser des éléments XML personnalisés. Le premier bloc de code montre comment incorporer des données XML dans le document. Il contient une liste de relecteurs, puis en utilisant les paramètres du classeur pour enregistrer le fichier XML`id` pour leur récupération future. Le deuxième bloc montre comment accéder à ce XML ultérieurement. Le paramètre « ContosoReviewXmlPartId » est chargé et transmis au classeur`customXmlParts`. Les données XML sont imprimées puis dans la console.

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
> `CustomXMLPart.namespaceUri` est renseigné uniquement si l’élément XML personnalisé niveau supérieur contient l’attribut`xmlns`.

## <a name="control-calculation-behavior"></a>Contrôler le comportement de calcul

### <a name="set-calculation-mode"></a>Définir le mode de calcul

Par défaut, Excel recalcule les résultats d’une formule chaque fois qu’une cellule référencée est modifiée. Le performances de votre complément peuvent profiter de l’ajustement de ce comportement de calcul. L’objet Application a une `calculationMode` propriété de type `CalculationMode`. Elle peut être définie sur les valeurs suivantes.

- `automatic`: Le comportement de recalcul par défaut dans lequel Excel calcule les résultats d’une nouvelle formule chaque fois que les données pertinentes sont modifiées.
- `automaticExceptTables`: Identique `automatic`, sauf que les modifications apportées à des valeurs dans les tableaux sont ignorées.
- `manual`: Calculs sont uniquement effectués lorsque l’utilisateur ou un complément les demande.

### <a name="set-calculation-type"></a>Définir le type de calcul

L’objet [Application](/javascript/api/excel/excel.application) fournit une méthode pour forcer un nouveau calcul immédiat. `Application.calculate(calculationType)` démarre un recalcul manuel basé sur la valeur `calculationType`. Les valeurs suivantes peuvent être spécifiées.

- `full`: Recalculer toutes les formules dans tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.
- `fullRebuild`: Revérifier les formules dépendantes, puis recalculer toutes les formules de tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.
- `recalculate`: Recalculer des formules qui ont changé (ou marqués par programme pour le recalcul) depuis le dernier calcul et les formules dépendantes, dans tous les classeurs actifs.

> [!NOTE]
> Pour plus d’informations sur le recalcul, voir l’article [recalcul de modification, l’itération ou la précision](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Suspendre temporairement les calculs

L’API Excel vous permet également de désactiver les compléments calculs jusqu'à ce que `RequestContext.sync()` ne soit appelé. Cette opération est effectuée avec `suspendApiCalculationUntilNextSync()`. Utilisez cette méthode lorsque votre complément modifie de grandes plages sans avoir à accéder aux données entre les modifications.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation"></a>Détecter l’activation d’unbook

Votre add-in peut détecter lorsqu’un workbook est activé. Un workbook devient *inactif* lorsque l’utilisateur bascule le focus vers un autre workbook, vers une autre application ou (dans Excel sur le Web) vers un autre onglet du navigateur web. Un workbook est *activé lorsque* l’utilisateur renvoie le focus au workbook. L’activation du workbook peut déclencher des fonctions de rappel dans votre complément, telles que l’actualisation des données du workbook.

Pour détecter lorsqu’un workbook est activé, inscrivez un [handler](excel-add-ins-events.md#register-an-event-handler) d’événements pour l’événement [onActivated](/javascript/api/excel/excel.workbook#onActivated) d’un workbook. Les handlers d’événements de l’événement reçoivent un `onActivated` [objet WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) lorsque l’événement se déclenche.

> [!IMPORTANT]
> `onActivated`L’événement ne détecte pas lorsqu’un workbook est ouvert. Cet événement détecte uniquement lorsqu’un utilisateur bascule le focus vers un workbook déjà ouvert.

L’exemple de code suivant montre comment inscrire le handler d’événements et `onActivated` configurer une fonction de rappel.

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

## <a name="save-the-workbook"></a>Enregistrer le classeur

`Workbook.save` enregistre le classeur dans un espace de stockage permanent. La `save` méthode prend un seul paramètre facultatif qui peut être `saveBehavior` l’une des valeurs suivantes.

- `Excel.SaveBehavior.save` (par défaut) : le fichier est enregistré sans inviter l’utilisateur à spécifier le nom de fichier et l’emplacement d’enregistrement. Si le fichier n’a pas été enregistré précédemment, il est enregistré dans l’emplacement par défaut. Si le fichier a été enregistré précédemment, il est enregistré au même emplacement.
- `Excel.SaveBehavior.prompt` : si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement. Si le fichier a été enregistré précédemment, il est enregistré dans le même emplacement et l’utilisateur ne reçoit pas d’invite.

> [!CAUTION]
> Si l’utilisateur est invité à enregistrer mais annule alors l’opération, `save` renvoie une erreur.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>Fermer le classeur

`Workbook.close` ferme le classeur, ainsi que des compléments qui sont associées au classeur (l’application Excel reste ouverte). La `close` méthode prend un seul paramètre facultatif qui peut être `closeBehavior` l’une des valeurs suivantes.

- `Excel.CloseBehavior.save` (par défaut) : le fichier est enregistré avant d’être fermé. Si le fichier n’a pas été enregistré précédemment, l’utilisateur sera invité à spécifier le nom de fichier et l’emplacement d’enregistrement.
- `Excel.CloseBehavior.skipSave` : le fichier est fermé immédiatement, sans enregistrer. Les modifications non enregistrées sont perdues.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser les feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md)
