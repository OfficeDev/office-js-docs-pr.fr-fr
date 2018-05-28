---
title: Conversion d?un projet de compl?ment Office dans Visual Studio au format TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 05e845b9d085b64b0534d28053dcd5ca3c7b403e
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/25/2018
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Conversion d?un projet de compl?ment Office dans Visual Studio au format TypeScript

Vous pouvez utiliser le mod?le de compl?ment Office dans Visual Studio pour cr?er un compl?ment qui utilise JavaScript, puis convertir le projet de compl?ment au format TypeScript. En utilisant Visual Studio pour cr?er le projet compl?ment, vous ne devez pas cr?er votre projet TypeScript de compl?ment Office ? partir de z?ro. 

Cet article explique comment cr?er un compl?ment Excel ? l?aide de Visual Studio et convertir le projet de compl?ment de JavaScript au format TypeScript. Vous pouvez suivre la m?me proc?dure pour convertir d'autres types de projet JavaScript de compl?ment Office au format TypeScript dans Visual Studio.

> [!NOTE]
> Pour cr?er un projet TypeScript de compl?ment Office sans utiliser Visual Studio, suivez les instructions de la section "N'importe quel ?diteur" d'un [D?marrage rapide en 5 minutes](../index.yml) et choisissez `TypeScript` ? l'invite du [G?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office).

## <a name="prerequisites"></a>Conditions pr?alables

- [Visual Studio 2017](https://www.visualstudio.com/vs/) avec la charge de travail de **d?veloppement Office/SharePoint** install?e

    > [!NOTE]
    > Si vous avez d?j? install? Visual Studio 2017, [utilisez Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) pour v?rifier que la charge de travail de **d?veloppement Office/SharePoint** est bien install?e. 

- TypeScript 2.3 pour Visual Studio 2017

    > [!NOTE]
    > TypeScript doit ?tre install? par d?faut avec Visual Studio 2017, mais vous pouvez [utiliser le programme d?installation Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) pour confirmer qu?il est install?. Dans le programme d?installation Visual Studio, s?lectionnez l?onglet **Composants individuels**, puis v?rifiez que l?option **SDK TypeScript 2.3** est s?lectionn?e sous **SDK, biblioth?ques et frameworks**.

- Excel 2016

## <a name="create-the-add-in-project"></a>Cr?ation du projet de compl?ment

1. Ouvrez Visual Studio, puis sur la barre de menus Visual Studio, s?lectionnez **Fichier** > **Nouveau** > **Projet**.

2. Dans la liste des types de projet sous **Visual C#** ou **Visual Basic**, d?veloppez **Office/SharePoint**, choisissez **Compl?ments**, puis **Compl?ment web Excel** pour le type de projet. 

3. Nommez le projet, puis cliquez sur **OK**.

4. Dans la fen?tre de dialogue **Cr?er un compl?ment Office**, s?lectionnez **Ajouter de nouvelles fonctionnalit?s ? Excel**, puis s?lectionnez **Terminer** pour cr?er le projet.

5. Visual Studio cr?e une solution et ses deux projets apparaissent dans l?**explorateur de solutions**. Le fichier **Home.html** s?ouvre dans Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Convertir le projet de compl?ment au format TypeScript

1. Dans l?**Explorateur de solutions**, renommez le fichier **Home.js** comme suit : **Home.ts**.

    > [!NOTE]
    > Dans votre projet TypeScript, vous pouvez avoir un m?lange de fichiers TypeScript et JavaScript, qui seront compil?s. En effet, TypeScript est un sur-ensemble typ? de code JavaScript compil? en code JavaScript. 

2. S?lectionnez **Oui** lorsque vous ?tes invit? ? confirmer la modification de l?extension du nom de fichier.

3. Cr?ez un fichier nomm? **Office.d.ts** dans la racine du projet d?application web.

4. Dans un navigateur web, ouvrez le [fichier de d?finitions de types pour Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copiez le contenu de ce fichier dans le presse-papiers.

5. Dans Visual Studio, ouvrez le fichier **Office.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.

6. Cr?ez un fichier nomm? **jQuery.d.ts** dans la racine du projet d?application web.

7. Dans un navigateur web, ouvrez le [fichier de d?finitions de types pour jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts). Copiez le contenu de ce fichier dans le presse-papiers.

8. Dans Visual Studio, ouvrez le fichier **jQuery.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.

9. Dans Visual Studio, cr?ez un fichier nomm? **tsconfig.json** dans la racine du projet d?application web.

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

## <a name="run-the-converted-add-in-project"></a>Ex?cuter le projet de compl?ment converti

1. Dans Visual Studio, appuyez sur F5 ou s?lectionnez le bouton **D?marrer** pour lancer Excel avec le bouton du compl?ment **Show Taskpane** (Afficher le volet Office) affich? dans le ruban. Le compl?ment sera h?berg? localement sur IIS.

2. Dans Excel, s?lectionnez l?onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du compl?ment.

3. Dans la feuille de calcul, s?lectionnez les neuf cellules qui contiennent des nombres.

4. Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage s?lectionn?e contenant la valeur la plus ?lev?e.

## <a name="homets-code-file"></a>Fichier de code Home.ts

Par exemple, l?extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications d?crites pr?c?demment ont ?t? appliqu?es. Ce code inclut le nombre minimal de modifications n?cessaires afin que votre compl?ment fonctionne.

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
            
            // If not using Excel 2016, use fallback logic.
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

* [Discussion sur la mise en ?uvre de promesses sur StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Exemples de compl?ments Office sur GitHub](https://github.com/officedev)
