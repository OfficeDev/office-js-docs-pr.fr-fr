---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 894cdcb8360a26dfb0f2d5ddbf06cbd7c52d5623
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945347"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Conversion d’un projet de complément Office dans Visual Studio au format TypeScript

Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript. En utilisant Visual Studio pour créer le projet complément, vous ne devez pas créer votre projet TypeScript de complément Office à partir de zéro. 

Cet article explique comment créer un complément Excel à l’aide de Visual Studio et convertir le projet de complément de JavaScript au format TypeScript. Vous pouvez utiliser le même processus pour convertir les autres types de projet JavaScript de complément Office au format TypeScript dans Visual Studio.

> [!NOTE]
> Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section « Tous les éditeurs » d’un [démarrage rapide en 5 minutes](../index.yml), puis sélectionnez `TypeScript` lorsque le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) vous y invite.

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio 2017](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée

    > [!NOTE]
    > Si vous avez déjà installé Visual Studio 2017, [utilisez Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée. 

- TypeScript 2.3 pour Visual Studio 2017

    > [!NOTE]
    > TypeScript doit être installé par défaut avec Visual Studio 2017, mais vous pouvez [utiliser le programme d’installation Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) pour confirmer qu’il est installé. Dans le programme d’installation Visual Studio, sélectionnez l’onglet **Composants individuels**, puis vérifiez que l’option **SDK TypeScript 2.3** est sélectionnée sous **SDK, bibliothèques et frameworks**.

- Excel 2016 ou version ultérieure

## <a name="create-the-add-in-project"></a>Création du projet de complément

1. Ouvrez Visual Studio, puis sur la barre de menus Visual Studio, sélectionnez **Fichier** > **Nouveau** > **Projet**.

2. Dans la liste des types de projet sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Excel** pour le type de projet. 

3. Nommez le projet, puis cliquez sur **OK**.

4. Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.

5. Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Convertir le projet de complément au format TypeScript

1. Dans l’**Explorateur de solutions**, renommez le fichier **Home.js** comme suit : **Home.ts**.

    > [!NOTE]
    > Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript. 

2. Sélectionnez **Oui** lorsque vous êtes invité à confirmer la modification de l’extension du nom de fichier.

3. Créez un fichier nommé **Office.d.ts** dans la racine du projet d’application web.

4. Dans un navigateur web, ouvrez le [fichier de définitions de types pour Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copiez le contenu de ce fichier dans le presse-papiers.

5. Dans Visual Studio, ouvrez le fichier **Office.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.

6. Créez un fichier nommé **jQuery.d.ts** dans la racine du projet d’application web.

7. Dans un navigateur web, ouvrez le [fichier de définitions de types pour jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts). Copiez le contenu de ce fichier dans le presse-papiers.

8. Dans Visual Studio, ouvrez le fichier **jQuery.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.

9. Dans Visual Studio, créez un fichier nommé **tsconfig.json** dans la racine du projet d’application web.

10. Ouvrez le fichier **tsconfig.json**, ajoutez le contenu suivant au fichier, puis enregistrez le fichier :

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. Ouvrez le fichier **Home.ts** et ajoutez les instructions suivantes en haut du fichier :

    ```javascript
    declare var fabric: any;
    ```

12. Dans le fichier **Home.ts**, remplacez **'1.1'** par **1.1** (autrement dit, supprimez les guillemets) dans la ligne suivante, puis enregistrez le fichier :

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a>Exécuter le projet de complément converti

1. Dans Visual Studio, appuyez sur F5 ou sélectionnez le bouton **Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** (Afficher le volet Office) affiché dans le ruban. Le complément sera hébergé localement sur IIS.

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

3. Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.

4. Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.

## <a name="homets-code-file"></a>Fichier de code Home.ts

Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées. Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016 or later, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```

## <a name="see-also"></a>Voir aussi

* [Discussion sur la mise en œuvre de promesses sur StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Exemples de compléments Office sur GitHub](https://github.com/officedev)
