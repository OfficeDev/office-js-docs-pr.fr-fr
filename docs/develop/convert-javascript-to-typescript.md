---
title: Conversion d’un projet de complément Office dans Visual Studio au format TypeScript
description: Découvrez comment convertir un projet de complément Office dans Visual Studio pour utiliser la machine à écrire.
ms.date: 09/01/2020
localization_priority: Normal
ms.openlocfilehash: 716222998666b014ab014298a4165678cf747165
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326295"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="a80dc-103">Conversion d’un projet de complément Office dans Visual Studio au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="a80dc-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="a80dc-104">Vous pouvez utiliser le modèle de complément Office dans Visual Studio pour créer un complément qui utilise JavaScript, puis convertir le projet de complément au format TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a80dc-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="a80dc-105">Cet article décrit ce processus de conversion pour un complément Excel.</span><span class="sxs-lookup"><span data-stu-id="a80dc-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="a80dc-106">Vous pouvez utiliser le même processus pour convertir les autres types de projet de complément Office de JavaScript au format TypeScript dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="a80dc-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a80dc-107">Cet article décrit les étapes *minimales* nécessaires pour garantir que, lorsque vous appuyez sur F5, le code est transpile en JavaScript qui est ensuite versions test chargées automatiquement dans Office.</span><span class="sxs-lookup"><span data-stu-id="a80dc-107">This article describes the *minimal* steps necessary to ensure that, when you press F5, the code will be transpiled to JavaScript which is then sideloaded automatically into Office.</span></span> <span data-ttu-id="a80dc-108">Toutefois, le code n’est pas très « TypeScripty ».</span><span class="sxs-lookup"><span data-stu-id="a80dc-108">However, the code is not very "TypeScripty".</span></span> <span data-ttu-id="a80dc-109">Par exemple, les variables sont déclarées avec le `var` mot clé au lieu de `let` et elles ne sont pas déclarées avec un type spécifié.</span><span class="sxs-lookup"><span data-stu-id="a80dc-109">For example, variables are declared with the `var` keyword instead of `let` and they are not declared with a specified type.</span></span> <span data-ttu-id="a80dc-110">Pour tirer pleinement parti du typage fort de la machine à écrire, pensez à apporter d’autres modifications au code.</span><span class="sxs-lookup"><span data-stu-id="a80dc-110">To take full advantage of the strong typing of TypeScript, consider making further changes to the code.</span></span> 

> [!NOTE]
> <span data-ttu-id="a80dc-111">Pour créer un projet TypeScript de complément Office sans utiliser Visual Studio, suivez les instructions de la section « Générateur Yeoman » d’un [démarrage rapide en 5 minutes](/office/dev/add-ins/), puis sélectionnez `TypeScript` quand le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) vous y invite.</span><span class="sxs-lookup"><span data-stu-id="a80dc-111">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](/office/dev/add-ins/) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a80dc-112">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="a80dc-112">Prerequisites</span></span>

- <span data-ttu-id="a80dc-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée</span><span class="sxs-lookup"><span data-stu-id="a80dc-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="a80dc-114">Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.</span><span class="sxs-lookup"><span data-stu-id="a80dc-114">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="a80dc-115">Si cette charge de travail n’est pas encore installée, utilisez Visual Studio Installer pour l’[installer](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="a80dc-115">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="a80dc-116">Kit de développement logiciel (SDK) TypeScript 2.3 ou version ultérieure (pour Visual Studio 2019)</span><span class="sxs-lookup"><span data-stu-id="a80dc-116">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="a80dc-117">Dans le [programme d’installation Visual Studio](/visualstudio/install/modify-visual-studio), sélectionnez l’onglet **Composants individuels**, puis faites défiler la page jusqu’à la section **SDK, bibliothèques et frameworks**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-117">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="a80dc-118">Dans cette section, vérifiez qu’au moins l’un des composants du **Kit de développement logiciel (SDK) TypeScript** (version 2.3 ou ultérieure) est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a80dc-118">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="a80dc-119">Si aucun des composants du **Kit de développement logiciel (SDK) TypeScript** n’est sélectionné, sélectionnez la dernière version disponible du SDK, puis sélectionnez le bouton **Modifier** pour [installer ce composant individuel](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="a80dc-119">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="a80dc-120">Excel 2016 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a80dc-120">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="a80dc-121">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="a80dc-121">Create the add-in project</span></span>

1. <span data-ttu-id="a80dc-122">Dans Visual Studio, choisissez **Créer un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-122">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="a80dc-123">À l’aide de la zone de recherche, entrez **complément**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-123">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="a80dc-124">Choisissez **Complément web Excel**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-124">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="a80dc-125">Nommez votre projet et sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-125">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="a80dc-126">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="a80dc-126">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="a80dc-p106">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="a80dc-p106">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="a80dc-129">Convertir le projet de complément au format TypeScript</span><span class="sxs-lookup"><span data-stu-id="a80dc-129">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="a80dc-130">Recherchez le fichier **Home.js** et renommez-le **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-130">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="a80dc-131">Recherchez le fichier **/Functions/FunctionFile.js** et renommez-le **FunctionFile.ts**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-131">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="a80dc-132">Recherchez le fichier **/Scripts/MessageBanner.js** et renommez-le **MessageBanner.ts**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-132">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="a80dc-133">Sous l’onglet **Outils**, choisissez **Gestionnaire de packages NuGet**, puis **Gérer un package NuGet pour Solution...**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-133">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="a80dc-134">L’onglet **Parcourir** étant sélectionné, entrez **jQuery. Machine à écrire. DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-134">With the **Browse** tab selected, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="a80dc-135">Installez ce package, ou mettez-le à jour s’il est déjà installé.</span><span class="sxs-lookup"><span data-stu-id="a80dc-135">Install this package, or update it if it's already installed.</span></span> <span data-ttu-id="a80dc-136">Cela permet de s’assurer que les définitions d’autodactylographiés jQuery sont incluses dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="a80dc-136">This will ensure the jQuery TypeScript definitions are included in your project.</span></span> <span data-ttu-id="a80dc-137">Les packages pour jQuery apparaissent dans un fichier généré par Visual Studio, appelé **packages.config**.</span><span class="sxs-lookup"><span data-stu-id="a80dc-137">The packages for jQuery appear in a file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a80dc-p108">Dans votre projet TypeScript, vous pouvez avoir un mélange de fichiers TypeScript et JavaScript, qui seront compilés. En effet, TypeScript est un sur-ensemble typé de code JavaScript compilé en code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a80dc-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

6. <span data-ttu-id="a80dc-140">Dans **Home.ts**, recherchez la ligne `Office.initialize = function (reason) {` et ajoutez une ligne immédiatement après celle-ci pour ajouter un polyfill à l’ensemble de `window.Promise`, comme illustré ici :</span><span class="sxs-lookup"><span data-stu-id="a80dc-140">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. <span data-ttu-id="a80dc-141">Dans **Home.ts**, recherchez la fonction `displaySelectedCells`, remplacez-la entièrement par le code suivant et enregistrez le fichier :</span><span class="sxs-lookup"><span data-stu-id="a80dc-141">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

8. <span data-ttu-id="a80dc-142">Dans **./Scripts/MessageBanner.ts**, recherchez la ligne `_onResize(null);` et remplacez-la par ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="a80dc-142">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="a80dc-143">Exécuter le projet de complément converti</span><span class="sxs-lookup"><span data-stu-id="a80dc-143">Run the converted add-in project</span></span>

1. <span data-ttu-id="a80dc-p109">Dans Visual Studio, appuyez sur**F5** ou sélectionnez le bouton**Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** affiché dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="a80dc-p109">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="a80dc-146">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="a80dc-146">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="a80dc-147">Dans la feuille de calcul, sélectionnez les neuf cellules qui contiennent des nombres.</span><span class="sxs-lookup"><span data-stu-id="a80dc-147">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="a80dc-148">Appuyez sur le bouton **Mettre en surbrillance** dans le volet Office pour mettre en surbrillance la cellule de la plage sélectionnée contenant la valeur la plus élevée.</span><span class="sxs-lookup"><span data-stu-id="a80dc-148">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="a80dc-149">Fichier de code Home.ts</span><span class="sxs-lookup"><span data-stu-id="a80dc-149">Home.ts code file</span></span>

<span data-ttu-id="a80dc-p110">Par exemple, l’extrait de code suivant affiche le contenu du fichier **Home.ts** une fois que les modifications décrites précédemment ont été appliquées. Ce code inclut le nombre minimal de modifications nécessaires afin que votre complément fonctionne.</span><span class="sxs-lookup"><span data-stu-id="a80dc-p110">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a80dc-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a80dc-152">See also</span></span>

- [<span data-ttu-id="a80dc-153">Discussion sur la mise en œuvre de promesses sur StackOverflow</span><span class="sxs-lookup"><span data-stu-id="a80dc-153">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="a80dc-154">Exemples de compléments Office sur GitHub</span><span class="sxs-lookup"><span data-stu-id="a80dc-154">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
