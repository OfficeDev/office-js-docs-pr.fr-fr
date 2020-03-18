---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: Découvrez comment convertir un projet de complément Office dans Visual Studio pour utiliser la machine à écrire.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 1dbb3503a521f1a7c3e71764a50f02708b667a11
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719041"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="f7994-103">Conversion d’un projet de complément Office dans Visual Studio au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="f7994-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="f7994-104">Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript.</span><span class="sxs-lookup"><span data-stu-id="f7994-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="f7994-105">Cet article décrit ce processus de conversion pour un complément Excel.</span><span class="sxs-lookup"><span data-stu-id="f7994-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="f7994-106">Vous pouvez utiliser le même processus pour convertir les autres types de projet de complément Office de JavaScript au format TypeScript dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="f7994-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="f7994-107">Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section « Générateur Yeoman » d’un [démarrage rapide en 5 minutes](../index.md), puis sélectionnez `TypeScript` quand le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) vous y invite.</span><span class="sxs-lookup"><span data-stu-id="f7994-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f7994-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="f7994-108">Prerequisites</span></span>

- <span data-ttu-id="f7994-109">[Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée</span><span class="sxs-lookup"><span data-stu-id="f7994-109">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="f7994-110">Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.</span><span class="sxs-lookup"><span data-stu-id="f7994-110">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="f7994-111">Si cette charge de travail n’est pas encore installée, utilisez Visual Studio Installer pour l’[installer](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="f7994-111">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="f7994-112">Kit de développement logiciel (SDK) TypeScript 2.3 ou version ultérieure (pour Visual Studio 2019)</span><span class="sxs-lookup"><span data-stu-id="f7994-112">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="f7994-113">Dans le [programme d’installation Visual Studio](/visualstudio/install/modify-visual-studio), sélectionnez l’onglet **Composants individuels**, puis faites défiler la page jusqu’à la section **SDK, bibliothèques et frameworks**.</span><span class="sxs-lookup"><span data-stu-id="f7994-113">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="f7994-114">Dans cette section, vérifiez qu’au moins l’un des composants du **Kit de développement logiciel (SDK) TypeScript** (version 2.3 ou ultérieure) est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f7994-114">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="f7994-115">Si aucun des composants du **Kit de développement logiciel (SDK) TypeScript** n’est sélectionné, sélectionnez la dernière version disponible du SDK, puis sélectionnez le bouton **Modifier** pour [installer ce composant individuel](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="f7994-115">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="f7994-116">Excel 2016 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="f7994-116">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="f7994-117">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="f7994-117">Create the add-in project</span></span>

1. <span data-ttu-id="f7994-118">Dans Visual Studio, choisissez **Créer un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="f7994-118">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="f7994-119">À l’aide de la zone de recherche, entrez **complément**.</span><span class="sxs-lookup"><span data-stu-id="f7994-119">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="f7994-120">Choisissez **Complément web Excel**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="f7994-120">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="f7994-121">Nommez votre projet et sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="f7994-121">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="f7994-122">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="f7994-122">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="f7994-p105">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="f7994-p105">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="f7994-125">Convertir le projet de complément au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="f7994-125">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="f7994-126">Recherchez le fichier **Home.js** et renommez-le **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="f7994-126">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="f7994-127">Recherchez le fichier **/Functions/FunctionFile.js** et renommez-le **FunctionFile.ts**.</span><span class="sxs-lookup"><span data-stu-id="f7994-127">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="f7994-128">Recherchez le fichier **/Scripts/MessageBanner.js** et renommez-le **MessageBanner.ts**.</span><span class="sxs-lookup"><span data-stu-id="f7994-128">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="f7994-129">Sous l’onglet **Outils**, choisissez **Gestionnaire de packages NuGet**, puis **Gérer un package NuGet pour Solution...**.</span><span class="sxs-lookup"><span data-stu-id="f7994-129">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="f7994-130">Avec l’onglet **Parcourir** sélectionné, entrez **office-js.TypeScript.DefinitelyTyped** dans la zone de recherche.</span><span class="sxs-lookup"><span data-stu-id="f7994-130">With the **Browse** tab selected, enter **office-js.TypeScript.DefinitelyTyped** into the search box.</span></span> <span data-ttu-id="f7994-131">Installer ou mettre à jour ce package s’il est déjà installé.</span><span class="sxs-lookup"><span data-stu-id="f7994-131">Install or update this package if it is already installed.</span></span> <span data-ttu-id="f7994-132">Cette opération ajoute les définitions de type TypeScript pour la bibliothèque Office.js à votre projet.</span><span class="sxs-lookup"><span data-stu-id="f7994-132">This will add the TypeScript type definitions for the Office.js library to your project.</span></span>

6. <span data-ttu-id="f7994-133">Dans la même zone de recherche, entrez **jquery.TypeScript.DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="f7994-133">In the same search box, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="f7994-134">Installer ou mettre à jour ce package s’il est déjà installé.</span><span class="sxs-lookup"><span data-stu-id="f7994-134">Install or update this package if it is already installed.</span></span> <span data-ttu-id="f7994-135">Cette opération permet d’ajouter les définitions de TypeScript jQuery dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="f7994-135">This will add the jQuery TypeScript definitions into your project.</span></span> <span data-ttu-id="f7994-136">Les packages pour jQuery et Office.js s’affichent désormais dans un nouveau fichier généré par Visual Studio, appelé **packages.config**.</span><span class="sxs-lookup"><span data-stu-id="f7994-136">The packages for both jQuery and Office.js will now appear in a new file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f7994-p108">Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f7994-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

7. <span data-ttu-id="f7994-139">Dans **Home.ts**, recherchez la ligne `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` et remplacez-la par ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="f7994-139">In **Home.ts**, find the line `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` and replace it with the following:</span></span>

    ```TypeScript
    if(!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
    ```

    > [!NOTE]
    > <span data-ttu-id="f7994-140">Pour l’instant, pour que le projet se compile correctement après avoir été converti en TypeScript, vous devez spécifier le numéro de l’ensemble de conditions requises sous forme de valeur numérique, comme illustré dans l’extrait de code précédent.</span><span class="sxs-lookup"><span data-stu-id="f7994-140">Currently, for the project to compile successfully after it's converted to TypeScript, you must specify the requirement set number as a numeric value as shown in the previous code snippet.</span></span> <span data-ttu-id="f7994-141">Malheureusement, cela signifie que vous ne pourrez pas utiliser `isSetSupported` pour tester la prise en charge de l’ensemble de conditions requises `1.10`, car la valeur numérique `1.10` a pour résultat `1.1` lors de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="f7994-141">Unfortunately this means you'll be unable to use `isSetSupported` to test for requirement set `1.10` support, as the numeric value `1.10` evaluates to `1.1` at runtime.</span></span> 
    > 
    > <span data-ttu-id="f7994-142">Ce problème est dû au fait que le package **office-js.TypeScript.DefinitelyTyped** NuGet est actuellement obsolète. Par conséquent, votre projet n’a pas accès aux dernières définitions TypeScript pour Office.js.</span><span class="sxs-lookup"><span data-stu-id="f7994-142">This problem is due to the **office-js.TypeScript.DefinitelyTyped** NuGet package currently being outdated, which means your project doesn't have access to the latest TypeScript definitions for Office.js.</span></span> <span data-ttu-id="f7994-143">Ce problème est en cours de traitement et cet article sera mis à jour une fois le problème résolu.</span><span class="sxs-lookup"><span data-stu-id="f7994-143">This issue is being addressed and this article will be updated when the issue is resolved.</span></span>

8. <span data-ttu-id="f7994-144">Dans **Home.ts**, recherchez la ligne `Office.initialize = function (reason) {` et ajoutez une ligne immédiatement après celle-ci pour ajouter un polyfill à l’ensemble de `window.Promise`, comme illustré ici :</span><span class="sxs-lookup"><span data-stu-id="f7994-144">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

9. <span data-ttu-id="f7994-145">Dans **Home.ts**, recherchez la fonction `displaySelectedCells`, remplacez-la entièrement par le code suivant et enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="f7994-145">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

10. <span data-ttu-id="f7994-146">Dans **./Scripts/MessageBanner.ts**, recherchez la ligne `_onResize(null);` et remplacez-la par ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="f7994-146">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="f7994-147">Exécuter le projet de complément converti</span><span class="sxs-lookup"><span data-stu-id="f7994-147">Run the converted add-in project</span></span>

1. <span data-ttu-id="f7994-p111">Dans Visual Studio, appuyez sur**F5** ou sélectionnez le bouton**Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** affiché dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="f7994-p111">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="f7994-150">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="f7994-150">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="f7994-151">Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.</span><span class="sxs-lookup"><span data-stu-id="f7994-151">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="f7994-152">Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.</span><span class="sxs-lookup"><span data-stu-id="f7994-152">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="f7994-153">Fichier de code Home.ts</span><span class="sxs-lookup"><span data-stu-id="f7994-153">Home.ts code file</span></span>

<span data-ttu-id="f7994-p112">Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées. Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.</span><span class="sxs-lookup"><span data-stu-id="f7994-p112">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f7994-156">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f7994-156">See also</span></span>

- [<span data-ttu-id="f7994-157">Discussion sur la mise en œuvre de promesses sur StackOverflow</span><span class="sxs-lookup"><span data-stu-id="f7994-157">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="f7994-158">Exemples de compléments Office sur GitHub</span><span class="sxs-lookup"><span data-stu-id="f7994-158">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
