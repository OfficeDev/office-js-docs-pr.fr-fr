---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: ''
ms.date: 10/30/2018
ms.openlocfilehash: d2a092cb48864cb9a4c9e791e3485963d0329ed2
ms.sourcegitcommit: 161a0625646a8c2ebaf1773c6369ee7cc96aa07b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/01/2018
ms.locfileid: "25891801"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="d9530-102">Conversion d’un projet de complément Office dans Visual Studio au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="d9530-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="d9530-103">Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d9530-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="d9530-104">Cet article décrit ce processus de conversion pour un complément Excel.</span><span class="sxs-lookup"><span data-stu-id="d9530-104">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="d9530-105">Vous pouvez utiliser le même processus pour convertir les autres types de projet de complément Office de JavaScript au format TypeScript dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d9530-105">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="d9530-106">Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section « Tous les éditeurs » d’un [démarrage rapide en 5 minutes](../index.yml), puis sélectionnez `TypeScript` quand le [générateur Yeoman pour les compléments Office](https://github.com/officedev/generator-office) vous y invite.</span><span class="sxs-lookup"><span data-stu-id="d9530-106">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d9530-107">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="d9530-107">Prerequisites</span></span>

- <span data-ttu-id="d9530-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée</span><span class="sxs-lookup"><span data-stu-id="d9530-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="d9530-109">Si vous avez déjà installé Visual Studio 2017, [utilisez Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.</span><span class="sxs-lookup"><span data-stu-id="d9530-109">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="d9530-110">Si cette charge de travail n’est pas encore installée, utilisez Visual Studio Installer pour l’[installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="d9530-110">If this workload is not yet installed, use the Visual Studio Installer to [install it](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span></span>

- <span data-ttu-id="d9530-111">Kit de développement logiciel (SDK) TypeScript 2.3 ou version ultérieure (pour Visual Studio 2017)</span><span class="sxs-lookup"><span data-stu-id="d9530-111">TypeScript SDK version 2.3 or later (for Visual Studio 2017)</span></span>

    > [!TIP]
    > <span data-ttu-id="d9530-112">Dans le [programme d’installation Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio), sélectionnez l’onglet **Composants individuels**, puis faites défiler la page jusqu’à la section **SDK, bibliothèques et frameworks**.</span><span class="sxs-lookup"><span data-stu-id="d9530-112">In the [Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="d9530-113">Dans cette section, vérifiez qu’au moins l’un des composants du **Kit de développement logiciel (SDK) TypeScript** (version 2.3 ou ultérieure) est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d9530-113">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="d9530-114">Si aucun des composants du **Kit de développement logiciel (SDK) TypeScript** n’est sélectionné, sélectionnez la dernière version disponible du SDK, puis sélectionnez le bouton **Modifier** pour [installer ce composant individuel](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="d9530-114">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span></span> 

- <span data-ttu-id="d9530-115">Excel 2016 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="d9530-115">Excel 2016, version 6769.2011 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="d9530-116">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="d9530-116">Create the add-in project</span></span>

1. <span data-ttu-id="d9530-117">Ouvrez Visual Studio, puis sur la barre de menus Visual Studio, sélectionnez **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="d9530-117">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="d9530-118">Dans la liste des types de projet sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Excel** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="d9530-118">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="d9530-119">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="d9530-119">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="d9530-120">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="d9530-120">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="d9530-p104">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d9530-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="d9530-123">Convertir le projet de complément au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="d9530-123">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="d9530-124">Dans l’**Explorateur de solutions**, renommez le fichier **Home.js** comme suit : **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="d9530-124">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="d9530-p105">Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d9530-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="d9530-127">Sélectionnez **Oui** lorsque vous êtes invité à confirmer la modification de l’extension du nom de fichier.</span><span class="sxs-lookup"><span data-stu-id="d9530-127">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="d9530-128">Créez un fichier nommé **Office.d.ts** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="d9530-128">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="d9530-129">Dans un navigateur web, ouvrez le [fichier de définitions de types pour Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="d9530-129">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="d9530-130">Copiez le contenu de ce fichier dans le presse-papiers.</span><span class="sxs-lookup"><span data-stu-id="d9530-130">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="d9530-131">Dans Visual Studio, ouvrez le fichier **Office.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d9530-131">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="d9530-132">Créez un fichier nommé **jQuery.d.ts** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="d9530-132">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="d9530-133">Dans un navigateur web, ouvrez le [fichier de définitions de types pour jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts).</span><span class="sxs-lookup"><span data-stu-id="d9530-133">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts).</span></span> <span data-ttu-id="d9530-134">Copiez le contenu de ce fichier dans le presse-papiers.</span><span class="sxs-lookup"><span data-stu-id="d9530-134">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="d9530-135">Dans Visual Studio, ouvrez le fichier **jQuery.d.ts**, collez le contenu du presse-papiers dans le fichier, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d9530-135">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="d9530-136">Dans Visual Studio, créez un fichier nommé **tsconfig.json** dans la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="d9530-136">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="d9530-137">Ouvrez le fichier **tsconfig.json**, ajoutez le contenu suivant au fichier, puis enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="d9530-137">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```json
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="d9530-138">Ouvrez le fichier **Home.ts** et ajoutez les instructions suivantes en haut du fichier :</span><span class="sxs-lookup"><span data-stu-id="d9530-138">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```typescript
    declare var fabric: any;
    ```

12. <span data-ttu-id="d9530-139">Dans le fichier **Home.ts**, remplacez **'1.1'** par **1.1** (autrement dit, supprimez les guillemets) dans la ligne suivante :</span><span class="sxs-lookup"><span data-stu-id="d9530-139">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```typescript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

13. <span data-ttu-id="d9530-140">Dans le fichier **Home.ts**, recherchez la fonction `displaySelectedCells`, remplacez-la par le code suivant, puis enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="d9530-140">In the **Home.ts** file, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

    ```typescript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
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

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="d9530-141">Exécuter le projet de complément converti</span><span class="sxs-lookup"><span data-stu-id="d9530-141">Run the converted add-in project</span></span>

1. <span data-ttu-id="d9530-p108">Dans Visual Studio, appuyez sur F5 ou sélectionnez le bouton **Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** (Afficher le volet Office) affiché dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="d9530-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="d9530-144">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d9530-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="d9530-145">Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.</span><span class="sxs-lookup"><span data-stu-id="d9530-145">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="d9530-146">Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.</span><span class="sxs-lookup"><span data-stu-id="d9530-146">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="d9530-147">Fichier de code Home.ts</span><span class="sxs-lookup"><span data-stu-id="d9530-147">Home.ts code file</span></span>

<span data-ttu-id="d9530-148">Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées.</span><span class="sxs-lookup"><span data-stu-id="d9530-148">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="d9530-149">Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.</span><span class="sxs-lookup"><span data-stu-id="d9530-149">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
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

## <a name="see-also"></a><span data-ttu-id="d9530-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d9530-150">See also</span></span>

* [<span data-ttu-id="d9530-151">Discussion sur la mise en œuvre de promesses sur StackOverflow</span><span class="sxs-lookup"><span data-stu-id="d9530-151">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="d9530-152">Exemples de compléments Office sur GitHub</span><span class="sxs-lookup"><span data-stu-id="d9530-152">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
