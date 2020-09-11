---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: Découvrez comment convertir un projet de complément Office dans Visual Studio pour utiliser la machine à écrire.
ms.date: 09/01/2020
localization_priority: Normal
ms.openlocfilehash: e05861e3fef79f87afc820eb62b2a52aaa953f31
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430484"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Conversion d’un projet de complément Office dans Visual Studio au format TypeScript

Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript. Cet article décrit ce processus de conversion pour un complément Excel. Vous pouvez utiliser le même processus pour convertir les autres types de projet de complément Office de JavaScript au format TypeScript dans Visual Studio.

> [!IMPORTANT]
> Cet article décrit les étapes *minimales* nécessaires pour garantir que, lorsque vous appuyez sur F5, le code est transpile en JavaScript qui est ensuite versions test chargées automatiquement dans Office. Toutefois, le code n’est pas très « TypeScripty ». Par exemple, les variables sont déclarées avec le `var` mot clé au lieu de `let` et elles ne sont pas déclarées avec un type spécifié. Pour tirer pleinement parti du typage fort de la machine à écrire, pensez à apporter d’autres modifications au code. 

> [!NOTE]
> Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section « Générateur Yeoman » d’un [démarrage rapide en 5 minutes](/office/dev/add-ins/), puis sélectionnez `TypeScript` quand le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) vous y invite.

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée

    > [!TIP]
    > Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée. Si cette charge de travail n’est pas encore installée, utilisez Visual Studio Installer pour l’[installer](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads).

- Kit de développement logiciel (SDK) TypeScript 2.3 ou version ultérieure (pour Visual Studio 2019)

    > [!TIP]
    > Dans le [programme d’installation Visual Studio](/visualstudio/install/modify-visual-studio), sélectionnez l’onglet **Composants individuels**, puis faites défiler la page jusqu’à la section **SDK, bibliothèques et frameworks**. Dans cette section, vérifiez qu’au moins l’un des composants du **Kit de développement logiciel (SDK) TypeScript** (version 2.3 ou ultérieure) est sélectionné. Si aucun des composants du **Kit de développement logiciel (SDK) TypeScript** n’est sélectionné, sélectionnez la dernière version disponible du SDK, puis sélectionnez le bouton **Modifier** pour [installer ce composant individuel](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components). 

- Excel 2016 ou version ultérieure

## <a name="create-the-add-in-project"></a>Création du projet de complément

1. Dans Visual Studio, choisissez **Créer un nouveau projet**.

2. À l’aide de la zone de recherche, entrez **complément**. Choisissez **Complément web Excel**, puis sélectionnez **Suivant**.

3. Nommez votre projet et sélectionnez **Créer**.

4. Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.

5. Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Convertir le projet de complément au format TypeScript

1. Recherchez le fichier **Home.js** et renommez-le **Home.ts**.

2. Recherchez le fichier **/Functions/FunctionFile.js** et renommez-le **FunctionFile.ts**.

3. Recherchez le fichier **/Scripts/MessageBanner.js** et renommez-le **MessageBanner.ts**.

4. Sous l’onglet **Outils**, choisissez **Gestionnaire de packages NuGet**, puis **Gérer un package NuGet pour Solution...**.

5. L’onglet **Parcourir** étant sélectionné, entrez **jQuery. Machine à écrire. DefinitelyTyped**. Installez ce package, ou mettez-le à jour s’il est déjà installé. Cela permet de s’assurer que les définitions d’autodactylographiés jQuery sont incluses dans votre projet. Les packages pour jQuery apparaissent dans un fichier généré par Visual Studio, appelé **packages.config**.

    > [!NOTE]
    > Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript.

6. Dans **Home.ts**, recherchez la ligne `Office.initialize = function (reason) {` et ajoutez une ligne immédiatement après celle-ci pour ajouter un polyfill à l’ensemble de `window.Promise`, comme illustré ici :

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. Dans **Home.ts**, recherchez la fonction `displaySelectedCells`, remplacez-la entièrement par le code suivant et enregistrez le fichier :

    ```TypeScript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }
    ```

8. Dans **./Scripts/MessageBanner.ts**, recherchez la ligne `_onResize(null);` et remplacez-la par ce qui suit :

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a>Exécuter le projet de complément converti

1. Dans Visual Studio, appuyez sur**F5** ou sélectionnez le bouton**Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** affiché dans le ruban. Le complément sera hébergé localement sur IIS.

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

3. Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.

4. Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.

## <a name="homets-code-file"></a>Fichier de code Home.ts

Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées. Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.

```typescript
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If you're using Excel 2013, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
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
            $('#highlight-button').click(highlightHighestValue);
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

    function highlightHighestValue() {
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
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
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

- [Discussion sur la mise en œuvre de promesses sur StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [Exemples de compléments Office sur GitHub](https://github.com/officedev)
