---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 05e845b9d085b64b0534d28053dcd5ca3c7b403e
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/25/2018
ms.locfileid: "19476528"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="12a4c-102">Conversion d’un projet de complément Office dans Visual Studio au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="12a4c-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="12a4c-103">Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript.</span><span class="sxs-lookup"><span data-stu-id="12a4c-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="12a4c-104">En utilisant Visual Studio pour créer le projet complément, vous ne devez pas créer votre projet TypeScript de complément Office à partir de zéro.</span><span class="sxs-lookup"><span data-stu-id="12a4c-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="12a4c-105">Cet article explique comment créer un complément Excel à l’aide de Visual Studio et convertir le projet de complément de JavaScript au format TypeScript.</span><span class="sxs-lookup"><span data-stu-id="12a4c-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="12a4c-106">Vous pouvez suivre la même procédure pour convertir d'autres types de projet JavaScript de complément Office au format TypeScript dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="12a4c-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="12a4c-107">Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section "N'importe quel éditeur" d'un [Démarrage rapide en 5 minutes](../index.yml) et choisissez `TypeScript` à l'invite du [Générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="12a4c-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="12a4c-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="12a4c-108">Prerequisites</span></span>

- <span data-ttu-id="12a4c-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée</span><span class="sxs-lookup"><span data-stu-id="12a4c-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="12a4c-110">Si vous avez déjà installé Visual Studio 2017, [utilisez Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.</span><span class="sxs-lookup"><span data-stu-id="12a4c-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="12a4c-111">TypeScript 2.3 pour Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="12a4c-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="12a4c-112">TypeScript doit être installé par défaut avec Visual Studio 2017, mais vous pouvez [utiliser le programme d’installation Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) pour confirmer qu’il est installé.</span><span class="sxs-lookup"><span data-stu-id="12a4c-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="12a4c-113">Dans le programme d’installation Visual Studio, sélectionnez l’onglet **Composants individuels**, puis vérifiez que l’option **SDK TypeScript 2.3** est sélectionnée sous **SDK, bibliothèques et frameworks**.</span><span class="sxs-lookup"><span data-stu-id="12a4c-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="12a4c-114">Excel 2016</span><span class="sxs-lookup"><span data-stu-id="12a4c-114">Excel 2016</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="12a4c-115">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="12a4c-115">Create the add-in project</span></span>

1. <span data-ttu-id="12a4c-116">Ouvrez Visual Studio, puis sur la barre de menus Visual Studio, sélectionnez **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="12a4c-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="12a4c-117">Dans la liste des types de projet sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Excel** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="12a4c-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="12a4c-118">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="12a4c-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="12a4c-119">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="12a4c-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="12a4c-p104">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="12a4c-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="12a4c-122">Convertir le projet de complément au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="12a4c-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="12a4c-123">Dans l’**Explorateur de solutions**, renommez le fichier **Home.js** comme suit : **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="12a4c-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="12a4c-p105">Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="12a4c-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="12a4c-126">Sélectionnez **Oui** lorsque vous êtes invité à confirmer la modification de l’extension du nom de fichier.</span><span class="sxs-lookup"><span data-stu-id="12a4c-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="12a4c-127">Créez un fichier nommé **Office.d.ts** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="12a4c-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="12a4c-128">Dans un navigateur web, ouvrez le [fichier de définitions de types pour Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="12a4c-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="12a4c-129">Copiez le contenu de ce fichier dans le presse-papiers.</span><span class="sxs-lookup"><span data-stu-id="12a4c-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="12a4c-130">Dans Visual Studio, ouvrez le fichier **Office.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="12a4c-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="12a4c-131">Créez un fichier nommé **jQuery.d.ts** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="12a4c-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="12a4c-132">Dans un navigateur web, ouvrez le [fichier de définitions de types pour jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="12a4c-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="12a4c-133">Copiez le contenu de ce fichier dans le presse-papiers.</span><span class="sxs-lookup"><span data-stu-id="12a4c-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="12a4c-134">Dans Visual Studio, ouvrez le fichier **jQuery.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="12a4c-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="12a4c-135">Dans Visual Studio, créez un fichier nommé **tsconfig.json** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="12a4c-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="12a4c-136">Ouvrez le fichier **tsconfig.json**, ajoutez le contenu suivant au fichier, puis enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="12a4c-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="12a4c-137">Ouvrez le fichier **Home.ts** et ajoutez les instructions suivantes en haut du fichier :</span><span class="sxs-lookup"><span data-stu-id="12a4c-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="12a4c-138">Dans le fichier **Home.ts**, remplacez **'1.1'** par **1.1** (autrement dit, supprimez les guillemets) dans la ligne suivante, puis enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="12a4c-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="12a4c-139">Exécuter le projet de complément converti</span><span class="sxs-lookup"><span data-stu-id="12a4c-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="12a4c-p108">Dans Visual Studio, appuyez sur F5 ou sélectionnez le bouton **Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** (Afficher le volet Office) affiché dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="12a4c-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="12a4c-142">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="12a4c-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="12a4c-143">Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.</span><span class="sxs-lookup"><span data-stu-id="12a4c-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="12a4c-144">Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.</span><span class="sxs-lookup"><span data-stu-id="12a4c-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="12a4c-145">Fichier de code Home.ts</span><span class="sxs-lookup"><span data-stu-id="12a4c-145">Home.ts code file</span></span>

<span data-ttu-id="12a4c-146">Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées.</span><span class="sxs-lookup"><span data-stu-id="12a4c-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="12a4c-147">Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.</span><span class="sxs-lookup"><span data-stu-id="12a4c-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="12a4c-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="12a4c-148">See also</span></span>

* [<span data-ttu-id="12a4c-149">Discussion sur la mise en œuvre de promesses sur StackOverflow</span><span class="sxs-lookup"><span data-stu-id="12a4c-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="12a4c-150">Exemples de compléments Office sur GitHub</span><span class="sxs-lookup"><span data-stu-id="12a4c-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
