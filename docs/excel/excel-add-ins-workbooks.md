---
title: Utiliser les classeurs utilisant l’API JavaScript Excel
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: 3d0cbc21d7e6b5c987df5a29d1aa83790c5685bc
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/22/2019
ms.locfileid: "30199591"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Utiliser les classeurs utilisant l’API JavaScript Excel

Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de classeurs utilisant l’API JavaScript pour Excel. Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Classeur**, reportez-vous à la rubrique [Objet classeur (API JavaScript pour Excel)](/javascript/api/excel/excel.workbook). Cet article décrit également les actions de niveau classeur effectuées via l’objet[Application](/javascript/api/excel/excel.application).

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

Vous pouvez accéder au classeur actif de votre complément en tant que chaîne codée en base 64 via [fichier découpage](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). La classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var workbookContents = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a>Insérer une copie d’un classeur existant dans l’offre actuelle

> [!NOTE]
> La fonction`WorksheetCollection.addFromBase64` est actuellement disponible uniquement en préversion publique. [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

L’exemple précédent montre un nouveau classeur créé à partir d’un classeur existant. Vous pouvez également copier la totalité ou une partie d’un classeur existant dans le tableau actuellement associé à votre complément. Un classeur[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) a la `addFromBase64`méthode pour insérer des copies de feuilles de calcul du classeur cible dans lui-même. Le fichier de l’autre classeur est passé en tant que chaîne codé en base 64, comme le `Excel.createWorkbook` appel.

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

L’exemple suivant montre des feuilles de calcul d’un classeur en cours d’insertion dans le classeur actif, juste après la feuille de calcul active. Notez que`null` est passé pour le`sheetNamesToInsert?: string[]` paramètre. Cela signifie que les feuilles de calcul sont insérées.

```js
var myFile = <HTMLInputElement>document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = (<string>(<FileReader>event.target).result).indexOf("base64,");
        var workbookContents = (<string>(<FileReader>event.target).result).substr(startIndex + 7);

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

La protection peut également être définie au niveau de la feuille de calcul pour empêcher la modification de données non souhaitée. Pour plus d’informations, voir la section**protection des données** de l’article[manipuler des feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md#data-protection).

> [!NOTE]
> Pour plus d’informations sur la protection du classeur dans Excel, voir l’article [protéger un classeur](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).

## <a name="access-document-properties"></a>Accès aux propriétés du document

Les objets classeur ont accès aux métadonnées de fichier Office, qui sont connues comme [propriétés du document](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). La propriété de l’objet classeur `properties` est un objet[DocumentProperties](/javascript/api/excel/excel.documentproperties) contenant ces valeurs de métadonnées. L’exemple de code suivant montre comment définir la propriété d’**auteur**.


```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

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
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
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

Par défaut, Excel recalcule les résultats d’une formule chaque fois qu’une cellule référencée est modifiée. Le performances de votre complément peuvent profiter de l’ajustement de ce comportement de calcul. L’objet Application a une `calculationMode` propriété de type `CalculationMode`. Peut être défini à l'aide des valeurs suivantes :


- `automatic`: Le comportement de recalcul par défaut dans lequel Excel calcule les résultats d’une nouvelle formule chaque fois que les données pertinentes sont modifiées.
- `automaticExceptTables`: Identique `automatic`, sauf que les modifications apportées à des valeurs dans les tableaux sont ignorées.
- `manual`: Calculs sont uniquement effectués lorsque l’utilisateur ou un complément les demande.

### <a name="set-calculation-type"></a>Définir le type de calcul

L’objet [Application](/javascript/api/excel/excel.application) fournit une méthode pour forcer un nouveau calcul immédiat. `Application.calculate(calculationType)` démarre un recalcul manuel basé sur la valeur `calculationType`. Les valeurs suivantes peuvent être utilisées :

- `full`: Recalculer toutes les formules dans tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.
- `fullRebuild`: Revérifier les formules dépendantes, puis recalculer toutes les formules de tous les classeurs ouverts, qu’elles aient changé depuis le dernier recalcul ou non.
- `recalculate`: Recalculer des formules qui ont changé (ou marqués par programme pour le recalcul) depuis le dernier calcul et les formules dépendantes, dans tous les classeurs actifs.

> [!NOTE]
> Pour plus d’informations sur le recalcul, voir l’article [recalcul de modification, l’itération ou la précision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Suspendre temporairement les calculs

L’API Excel vous permet également de désactiver les compléments calculs jusqu'à ce que `RequestContext.sync()` ne soit appelé. Cette opération est effectuée avec `suspendApiCalculationUntilNextSync()`. Utilisez cette méthode lorsque votre complément modifie de grandes plages sans avoir à accéder aux données entre les modifications.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Utiliser les feuilles de calcul à l’aide de l’API JavaScript Excel](excel-add-ins-worksheets.md)
- [Utiliser les plages à l’aide de l’API JavaScript Excel](excel-add-ins-ranges.md)
