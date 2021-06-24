---
title: Didacticiel sur le complément Excel
description: Dans ce didacticiel, vous allez développer un complément Excel qui crée, remplit, filtre et trie un tableau, crée un graphique, fige un en-tête de tableau, protège une feuille de calcul et ouvre une boîte de dialogue.
ms.date: 05/12/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: f169499e343d2fc7fac89f407b78717536add4fc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077238"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="eedea-103">Didacticiel : Créer un complément de volet de tâches de Excel</span><span class="sxs-lookup"><span data-stu-id="eedea-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="eedea-104">Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:</span><span class="sxs-lookup"><span data-stu-id="eedea-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
>
> - <span data-ttu-id="eedea-105">Crée un tableau</span><span class="sxs-lookup"><span data-stu-id="eedea-105">Creates a table</span></span>
> - <span data-ttu-id="eedea-106">Filtres et tris un tableau</span><span class="sxs-lookup"><span data-stu-id="eedea-106">Filters and sorts a table</span></span>
> - <span data-ttu-id="eedea-107">Crée un graphique (Chart)</span><span class="sxs-lookup"><span data-stu-id="eedea-107">Creates a chart</span></span>
> - <span data-ttu-id="eedea-108">Figer une en-tête de tableau</span><span class="sxs-lookup"><span data-stu-id="eedea-108">Freezes a table header</span></span>
> - <span data-ttu-id="eedea-109">Protège une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="eedea-109">Protects a worksheet</span></span>
> - <span data-ttu-id="eedea-110">Ouvrir une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="eedea-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="eedea-111">Si vous avez déjà exécuté le démarrage rapide [Créer votre premier complément du volet des tâches d’Excel](../quickstarts/excel-quickstart-jquery.md) à l’aide du générateur Yeoman et que vous souhaitez utiliser ce projet comme point de départ pour ce didacticiel, accédez directement à la section [Créer un tableau](#create-a-table) pour commencer ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="eedea-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="eedea-112">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eedea-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="eedea-113">Créer votre projet de complément</span><span class="sxs-lookup"><span data-stu-id="eedea-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="eedea-114">**Sélectionnez un type de projet :** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="eedea-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="eedea-115">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="eedea-115">**Choose a script type:** `JavaScript`</span></span>
- <span data-ttu-id="eedea-116">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="eedea-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="eedea-117">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="eedea-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Capture d’écran de l’interface de ligne de commande du générateur de compléments Yeoman Office.](../images/yo-office-excel.png)

<span data-ttu-id="eedea-119">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="eedea-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="eedea-120">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="eedea-120">Create a table</span></span>

<span data-ttu-id="eedea-121">Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.</span><span class="sxs-lookup"><span data-stu-id="eedea-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="eedea-122">Codage du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-122">Code the add-in</span></span>

1. <span data-ttu-id="eedea-123">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="eedea-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="eedea-124">Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="eedea-124">Open the file **./src/taskpane/taskpane.html**.</span></span>  <span data-ttu-id="eedea-125">Ce fichier contient la balise HTML du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="eedea-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="eedea-126">Recherchez l’élément `<main>` et supprimez toutes les lignes qui apparaissent après la balise `<main>` d’ouverture et avant la balise `</main>` de fermeture.</span><span class="sxs-lookup"><span data-stu-id="eedea-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="eedea-127">Ajoutez la balise suivante juste après la balise `<main>` d’ouverture :</span><span class="sxs-lookup"><span data-stu-id="eedea-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="eedea-128">Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="eedea-129">Ce fichier contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

6. <span data-ttu-id="eedea-130">Supprimez toutes les références au bouton `run` et à la fonction `run()` en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="eedea-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="eedea-131">Recherchez et supprimez la ligne `document.getElementById("run").onclick = run;`.</span><span class="sxs-lookup"><span data-stu-id="eedea-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="eedea-132">Recherchez et supprimez la fonction `run()` entière.</span><span class="sxs-lookup"><span data-stu-id="eedea-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="eedea-133">Au sein de l’appel de méthode `Office.onReady`, recherchez la ligne `if (info.host === Office.HostType.Excel) {` et ajoutez le code suivant immédiatement après cette ligne.</span><span class="sxs-lookup"><span data-stu-id="eedea-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="eedea-134">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-134">Note:</span></span>

    - <span data-ttu-id="eedea-p104">La première partie de ce code détermine si la version d' Excel de l'utilisateur prend en charge une version d'Excel.js qui inclut toutes les API que cette série de tutoriels utilisera. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l'interface utilisateur qui appellerait les API non prises en charge. Cela permettra à l'utilisateur de continuer à utiliser les parties du complément qui sont prises en charge par sa version d' Excel.</span><span class="sxs-lookup"><span data-stu-id="eedea-p104">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="eedea-138">La deuxième partie de ce code ajoute un gestionnaire d’événements pour le bouton `create-table`.</span><span class="sxs-lookup"><span data-stu-id="eedea-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="eedea-139">Ajoutez la fonction suivante à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="eedea-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="eedea-140">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-140">Note:</span></span>

    - <span data-ttu-id="eedea-p106">Votre logique métier Excel.js est ajoutée à la fonction qui est transmise à `Excel.run`. Cette logique n’est pas exécutée immédiatement. Au lieu de cela, elle est ajoutée à une file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="eedea-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="eedea-144">La méthode `context.sync` envoie toutes les commandes en file d’attente vers Excel pour exécution.</span><span class="sxs-lookup"><span data-stu-id="eedea-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="eedea-p107">L’élément `Excel.run` est suivi par un bloc `catch`. Il s’agit d’une meilleure pratique que vous devez toujours suivre.</span><span class="sxs-lookup"><span data-stu-id="eedea-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

9. <span data-ttu-id="eedea-147">À l’intérieur de la fonction `createTable()`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-147">Within the `createTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-148">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-148">Note:</span></span>

    - <span data-ttu-id="eedea-p109">Le code crée un tableau à l’aide de la méthode `add` de collection de tableau d’une feuille de calcul, qui existe toujours même si elle est vide. Il s’agit de la méthode standard de création d’objets Excel.js. Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel. Au lieu de cela, vous l’ajoutez à un objet de la collection parent.</span><span class="sxs-lookup"><span data-stu-id="eedea-p109">The code creates a table by using the `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="eedea-p110">Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau. La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes. Il s’agit d’un modèle courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.</span><span class="sxs-lookup"><span data-stu-id="eedea-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a common pattern, because the number of rows a table will have is often unknown when the table is created.</span></span>

    - <span data-ttu-id="eedea-156">Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eedea-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="eedea-157">À l’intérieur de la fonction `createTable()`, remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-157">Within the `createTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="eedea-158">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-158">Note:</span></span>

    - <span data-ttu-id="eedea-159">Les valeurs de cellule d’une plage sont définies avec un tableau de tableaux.</span><span class="sxs-lookup"><span data-stu-id="eedea-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="eedea-p112">Les nouvelles lignes sont créées dans un tableau en appelant la méthode `add` de collection de ligne du tableau. Vous pouvez ajouter plusieurs lignes dans un seul appel de `add` en incluant plusieurs tableaux de valeurs de cellule dans le tableau parent transmis en tant que deuxième paramètre.</span><span class="sxs-lookup"><span data-stu-id="eedea-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

    ```js
    expensesTable.getHeaderRowRange().values =
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ```

11. <span data-ttu-id="eedea-162">À l’intérieur de la fonction `createTable()`, remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-162">Within the `createTable()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="eedea-163">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-163">Note:</span></span>

    - <span data-ttu-id="eedea-164">Le code recherche une référence à la colonne **Amount** en transmettant son index de base zéro à la méthode `getItemAt` de collection de colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="eedea-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="eedea-165">Les objets de collection Excel.js, tels que `TableCollection`, `WorksheetCollection` et `TableColumnCollection` ont une propriété `items` qui correspond à un tableau de types d’objet enfant, comme `Table` ou `Worksheet` ou `TableColumn` ; mais un objet `*Collection` n’est pas lui-même un tableau.</span><span class="sxs-lookup"><span data-stu-id="eedea-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="eedea-166">Le code définit ensuite la plage de la colonne **Amount** sous la forme Euros à la deuxième décimale.</span><span class="sxs-lookup"><span data-stu-id="eedea-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span>

    - <span data-ttu-id="eedea-p114">Enfin, il s’assure que la largeur des colonnes et la hauteur des lignes sont assez grandes pour contenir l’élément de données le plus long (ou le plus haut). Notez que le code doit rechercher des objets `Range` à mettre en forme. Les objets `TableColumn` et `TableRow` n’ont pas de propriétés de mise en forme.</span><span class="sxs-lookup"><span data-stu-id="eedea-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="eedea-170">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-171">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-171">Test the add-in</span></span>

1. <span data-ttu-id="eedea-172">Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="eedea-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="eedea-173">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="eedea-173">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="eedea-174">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="eedea-174">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="eedea-175">Si vous testez votre complément sur Mac, exécutez la commande suivante dans le répertoire racine de votre projet avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="eedea-175">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="eedea-176">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="eedea-176">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="eedea-177">Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-177">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="eedea-178">Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Excel avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="eedea-178">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="eedea-179">Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-179">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="eedea-180">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="eedea-180">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="eedea-181">Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="eedea-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="eedea-182">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="eedea-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran du menu Accueil d’Excel avec le bouton Afficher le volet Office mis en évidence.](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="eedea-184">Dans le volet Office, sélectionnez le bouton **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="eedea-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Capture d’écran d’Excel, montrant un volet office de complément avec un bouton Créer un tableau et un tableau dans la feuille de calcul rempli de données Date, Commerçant, Catégorie et Montant.](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table&quot;></a><span data-ttu-id=&quot;eedea-186&quot;>Filtrer et trier un tableau</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-186&quot;>Filter and sort a table</span></span>

<span data-ttu-id=&quot;eedea-187&quot;>Dans cette étape du didacticiel, vous allez filtrer et trier le tableau que vous avez créé précédemment.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-187&quot;>In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name=&quot;filter-the-table&quot;></a><span data-ttu-id=&quot;eedea-188&quot;>Filtrage du tableau</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-188&quot;>Filter the table</span></span>

1. <span data-ttu-id=&quot;eedea-189&quot;>Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-189&quot;>Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id=&quot;eedea-190&quot;>Recherchez l’élément `<button>` du bouton `create-table`, puis ajoutez la balise suivante après cette ligne :</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-190&quot;>Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class=&quot;ms-Button&quot; id=&quot;filter-table&quot;>Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id=&quot;eedea-191&quot;>Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-191&quot;>Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id=&quot;eedea-192&quot;>Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `create-table`, puis ajoutez le code suivant après cette ligne :</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-192&quot;>Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById(&quot;filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="eedea-193">Ajoutez la fonction suivante à la fin du fichier :</span><span class="sxs-lookup"><span data-stu-id="eedea-193">Add the following function to the end of the file:</span></span>

    ```js
    function filterTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to filter out all expense categories except
            //        Groceries and Education.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="eedea-194">À l’intérieur de la fonction `filterTable()`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-194">Within the `filterTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-195">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-195">Note:</span></span>

   - <span data-ttu-id="eedea-p120">Le code obtient tout d’abord une référence à la colonne à filtrer en transférant le nom de la colonne à la méthode `getItem`, au lieu de transmettre son index à la méthode `getItemAt` comme le fait la méthode `createTable`. Puisque les utilisateurs peuvent déplacer des colonnes de tableau, la colonne d’un index donné peut être modifiée après la création du tableau. Par conséquent, il est préférable d’utiliser le nom de la colonne pour obtenir une référence de la colonne. Dans le didacticiel précédent, nous avons utilisé la méthode `getItemAt` en toute sécurité, car nous l’avons utilisée dans la même méthode que celle qui crée le tableau, il n’y a donc aucune chance qu’un utilisateur ait déplacé la colonne.</span><span class="sxs-lookup"><span data-stu-id="eedea-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="eedea-200">La méthode `applyValuesFilter` est l’une des nombreuses méthodes de filtrage sur l’objet `Filter`.</span><span class="sxs-lookup"><span data-stu-id="eedea-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ```

### <a name="sort-the-table"></a><span data-ttu-id="eedea-201">Tri du tableau</span><span class="sxs-lookup"><span data-stu-id="eedea-201">Sort the table</span></span>

1. <span data-ttu-id="eedea-202">Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="eedea-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="eedea-203">Recherchez l’élément `<button>` du bouton `filter-table`, puis ajoutez la balise suivante après cette ligne :</span><span class="sxs-lookup"><span data-stu-id="eedea-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="eedea-204">Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="eedea-205">Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `filter-table`, puis ajoutez le code suivant après cette ligne :</span><span class="sxs-lookup"><span data-stu-id="eedea-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="eedea-206">Ajoutez la fonction suivante à la fin du fichier :</span><span class="sxs-lookup"><span data-stu-id="eedea-206">Add the following function to the end of the file:</span></span>

    ```js
    function sortTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to sort the table by Merchant name.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="eedea-207">À l’intérieur de la fonction `sortTable()`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-207">Within the `sortTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-208">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-208">Note:</span></span>

   - <span data-ttu-id="eedea-209">Le code crée un tableau d’objets `SortField` qui ne comporte qu’un seul membre puisque le complément ne trie que la colonne Merchant.</span><span class="sxs-lookup"><span data-stu-id="eedea-209">The code creates an array of `SortField` objects, which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="eedea-210">La propriété `key` d’un objet `SortField` est l’index de la colonne utilisée pour le tri.</span><span class="sxs-lookup"><span data-stu-id="eedea-210">The `key` property of a `SortField` object is the zero-based index of the column used for sorting.</span></span> <span data-ttu-id="eedea-211">Les lignes du tableau sont triées sur la base des valeurs de la colonne référencée.</span><span class="sxs-lookup"><span data-stu-id="eedea-211">The rows of the table are sorted based on the values in the referenced column.</span></span>

   - <span data-ttu-id="eedea-212">Le membre `sort` d’un objet `Table` est un objet `TableSort`, et non une méthode.</span><span class="sxs-lookup"><span data-stu-id="eedea-212">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="eedea-213">Les objets `SortField` sont transmis à la méthode `apply` de l’objet `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="eedea-213">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

7. <span data-ttu-id="eedea-214">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-214">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-215">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-215">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="eedea-216">Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.</span><span class="sxs-lookup"><span data-stu-id="eedea-216">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="eedea-217">Si la tableau que vous avez ajoutée précédemment dans ce didacticiel ne figure pas dans la feuille de calcul ouverte, sélectionnez le bouton **Créer un tableau** dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-217">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="eedea-218">Choisissez le bouton **Filtrer le tableau** et le bouton **Trier le tableau** dans n’importe quel ordre.</span><span class="sxs-lookup"><span data-stu-id="eedea-218">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Capture d’écran d’Excel, avec les boutons Filtrer le tableau et Trier le tableau mis en évidence dans le volet Office Complément.](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a><span data-ttu-id="eedea-220">Création d’un graphique (chart)</span><span class="sxs-lookup"><span data-stu-id="eedea-220">Create a chart</span></span>

<span data-ttu-id="eedea-221">Dans cette étape du didacticiel, vous créerez un graphique à l’aide de données provenant du tableau précédemment créé, puis vous mettrez en forme le graphique.</span><span class="sxs-lookup"><span data-stu-id="eedea-221">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="eedea-222">Un graphique à l’aide de données du tableau de graphique (chart)</span><span class="sxs-lookup"><span data-stu-id="eedea-222">Chart a chart using table data</span></span>

1. <span data-ttu-id="eedea-223">Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="eedea-223">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="eedea-224">Recherchez l’élément `<button>` du bouton `sort-table`, puis ajoutez la balise suivante après cette ligne :</span><span class="sxs-lookup"><span data-stu-id="eedea-224">Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id="eedea-225">Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-225">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="eedea-226">Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `sort-table`, puis ajoutez le code suivant après cette ligne :</span><span class="sxs-lookup"><span data-stu-id="eedea-226">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="eedea-227">Ajoutez la fonction suivante à la fin du fichier :</span><span class="sxs-lookup"><span data-stu-id="eedea-227">Add the following function to the end of the file:</span></span>

    ```js
    function createChart() {
        Excel.run(function (context) {

            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="eedea-228">À l’intérieur de la fonction `createChart()`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-228">Within the `createChart()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-229">Pour exclure la ligne d’en-tête, le code utilise la méthode `Table.getDataBodyRange` pour obtenir la plage de données à représenter sous forme de graphique à la place de la méthode `getRange`.</span><span class="sxs-lookup"><span data-stu-id="eedea-229">Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="eedea-230">À l’intérieur de la fonction `createChart()`, remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-230">Within the `createChart()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="eedea-231">Notez les paramètres suivants :</span><span class="sxs-lookup"><span data-stu-id="eedea-231">Note the following parameters:</span></span>

   - <span data-ttu-id="eedea-p126">Le premier paramètre transmis à la méthode `add` spécifie le type de graphique. Il en existe plusieurs dizaines de types.</span><span class="sxs-lookup"><span data-stu-id="eedea-p126">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="eedea-234">Le deuxième paramètre spécifie la plage de données à inclure dans le graphique.</span><span class="sxs-lookup"><span data-stu-id="eedea-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="eedea-p127">Le troisième paramètre détermine si une série de points de données de la table doit être représentée par ligne ou par colonne. L’option `auto` indique à Excel de décider de la meilleure méthode.</span><span class="sxs-lookup"><span data-stu-id="eedea-p127">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise. The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');
    ```

8. <span data-ttu-id="eedea-237">À l’intérieur de la fonction `createChart()`, remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-237">Within the `createChart()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="eedea-238">La majeure partie du code est explicite.</span><span class="sxs-lookup"><span data-stu-id="eedea-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="eedea-239">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-239">Note:</span></span>

   - <span data-ttu-id="eedea-p129">Les paramètres de la méthode `setPosition` spécifient les cellules situées en haut à gauche et en bas à droite de la zone de feuille de calcul devant contenir le graphique. Excel peut ajuster des éléments, tels que la largeur de ligne pour que le graphique s’affiche correctement dans l’espace attribué.</span><span class="sxs-lookup"><span data-stu-id="eedea-p129">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>

   - <span data-ttu-id="eedea-p130">Une « série » est un ensemble de points de données dans une colonne du tableau. Étant donné qu’il n’existe qu’une seule colonne autre que de type chaîne dans le tableau, Excel déduit que la colonne est la seule colonne de points de données pour le graphique. Il interprète les autres colonnes comme des étiquettes de graphique. Par conséquent, il y aura simplement une série dans le graphique et un index 0. Il s’agit de celle à étiqueter avec « Valeur en &euro; ».</span><span class="sxs-lookup"><span data-stu-id="eedea-p130">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in &euro;".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in \u20AC';
    ```

9. <span data-ttu-id="eedea-247">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-247">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-248">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-248">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="eedea-249">Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.</span><span class="sxs-lookup"><span data-stu-id="eedea-249">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="eedea-250">Si la tableau que vous avez ajoutée précédemment dans ce didacticiel ne figure pas dans la feuille de calcul ouverte, sélectionnez le bouton **Créer un tableau**, puis le bouton **Filtrer un tableau** et le bouton **Trier un tableau** dans n’importe quel ordre..</span><span class="sxs-lookup"><span data-stu-id="eedea-250">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="eedea-p131">Sélectionnez le bouton **Créer un graphique**. Un graphique est créé dans lequel seules les données provenant des lignes filtrées sont incluses. Les étiquettes sur les points de données en bas sont organisées selon l’ordre de tri du graphique, à savoir les noms de marchand par ordre alphabétique inversé.</span><span class="sxs-lookup"><span data-stu-id="eedea-p131">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Capture d’écran d’Excel, avec un bouton Créer un graphique visible dans le volet Office du complément et un graphique dans la feuille de calcul affichant les données de dépenses d’alimentation et d’éducation.](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header&quot;></a><span data-ttu-id=&quot;eedea-255&quot;>Figer un en-tête de tableau</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-255&quot;>Freeze a table header</span></span>

<span data-ttu-id=&quot;eedea-p132&quot;>Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée. Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-p132&quot;>When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name=&quot;freeze-the-tables-header-row&quot;></a><span data-ttu-id=&quot;eedea-258&quot;>Figer la ligne d’en-tête du tableau</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-258&quot;>Freeze the table's header row</span></span>

1. <span data-ttu-id=&quot;eedea-259&quot;>Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-259&quot;>Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id=&quot;eedea-260&quot;>Recherchez l’élément `<button>` du bouton `create-chart`, puis ajoutez la balise suivante après cette ligne :</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-260&quot;>Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span>

    ```html
    <button class=&quot;ms-Button&quot; id=&quot;freeze-header&quot;>Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id=&quot;eedea-261&quot;>Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-261&quot;>Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id=&quot;eedea-262&quot;>Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `create-chart`, puis ajoutez le code suivant après cette ligne :</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;eedea-262&quot;>Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById(&quot;freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="eedea-263">Ajoutez la fonction suivante à la fin du fichier :</span><span class="sxs-lookup"><span data-stu-id="eedea-263">Add the following function to the end of the file:</span></span>

    ```js
    function freezeHeader() {
        Excel.run(function (context) {

            // TODO1: Queue commands to keep the header visible when the user scrolls.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="eedea-264">À l’intérieur de la fonction `freezeHeader()`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-264">Within the `freezeHeader()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-265">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-265">Note:</span></span>

   - <span data-ttu-id="eedea-266">La collection `Worksheet.freezePanes` est un ensemble de volets de la feuille de calcul qui sont épinglés, c’est-à-dire figés, lorsque vous faites défiler la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eedea-266">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="eedea-p134">La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être épinglées. `1` est transmis pour épingler la première rangée.</span><span class="sxs-lookup"><span data-stu-id="eedea-p134">The `freezeRows` method takes as a parameter the number of rows, from the top, that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="eedea-269">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-269">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-270">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-270">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="eedea-271">Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.</span><span class="sxs-lookup"><span data-stu-id="eedea-271">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="eedea-272">Si le tableau que vous avez ajouté précédemment dans ce didacticiel est présent dans la feuille de calcul, supprimez-le.</span><span class="sxs-lookup"><span data-stu-id="eedea-272">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="eedea-273">Dans le volet Office, sélectionnez le bouton **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="eedea-273">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="eedea-274">Dans le volet Office, sélectionnez le bouton **Figer l’en-tête**.</span><span class="sxs-lookup"><span data-stu-id="eedea-274">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="eedea-275">Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.</span><span class="sxs-lookup"><span data-stu-id="eedea-275">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Capture d’écran illustrant une feuille de calcul Excel avec un en-tête de tableau figé.](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="eedea-277">Protéger une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="eedea-277">Protect a worksheet</span></span>

<span data-ttu-id="eedea-278">Au cours de cette étape, vous allez ajouter un bouton au ruban pour activer ou désactiver la protection de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eedea-278">In this step of the tutorial, you'll add a button to the ribbon that toggles worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="eedea-279">Configuration du manifeste pour ajouter un deuxième bouton de ruban</span><span class="sxs-lookup"><span data-stu-id="eedea-279">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="eedea-280">Ouvrez le fichier manifeste **./manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="eedea-280">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="eedea-281">Recherchez l’élément `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="eedea-281">Locate the `<Control>` element.</span></span> <span data-ttu-id="eedea-282">Cet élément définit le bouton **Afficher le volet des pages** sur le ruban **Accueil** que vous utilisez pour lancer le complément.</span><span class="sxs-lookup"><span data-stu-id="eedea-282">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="eedea-283">Nous allons ajouter un deuxième bouton au même groupe sur le ruban **Accueil**.</span><span class="sxs-lookup"><span data-stu-id="eedea-283">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="eedea-284">Dans la balise de `</Control>` de fermeture et la balise de `</Group>` de fermeture, ajoutez la balise suivante.</span><span class="sxs-lookup"><span data-stu-id="eedea-284">In between the closing `</Control>` tag and the closing `</Group>` tag, add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="eedea-285">Dans le code XML que vous venez d’ajouter au fichier manifeste, remplacez `TODO1` par une chaîne qui attribue un ID unique au bouton au sein de ce fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="eedea-285">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="eedea-286">Étant donné que notre bouton va activer ou désactiver la protection de la feuille de calcul, utilisez « ToggleProtection ».</span><span class="sxs-lookup"><span data-stu-id="eedea-286">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="eedea-287">Lorsque vous avez terminé, l’étiquette d’ouverture de l’élément `Control` doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="eedea-287">When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="eedea-288">Les trois `TODO`s suivantes définissent les ID de ressource ou `resid`s.</span><span class="sxs-lookup"><span data-stu-id="eedea-288">The next three `TODO`s set resource IDs, or `resid`s.</span></span> <span data-ttu-id="eedea-289">Une ressource est une chaîne (d’une longueur maximale de 32 caractères). Vous allez créer ces trois chaînes lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="eedea-289">A resource is a string (with a maximum length of 32 characters), and you'll create these three strings in a later step.</span></span> <span data-ttu-id="eedea-290">Pour l’instant, vous devez attribuer des ID aux ressources.</span><span class="sxs-lookup"><span data-stu-id="eedea-290">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="eedea-291">L’étiquette du bouton doit indiquer « Toggle Protection », mais l’*ID* de cette chaîne doit être « ProtectionButtonLabel », donc l’élément `Label` doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="eedea-291">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="eedea-292">L’élément `SuperTip` définit l’info-bulle du bouton.</span><span class="sxs-lookup"><span data-stu-id="eedea-292">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="eedea-293">Le titre de l’info-bulle doit être identique à l’étiquette du bouton, nous utilisons donc le même ID de ressource : « ProtectionButtonLabel ».</span><span class="sxs-lookup"><span data-stu-id="eedea-293">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="eedea-294">La description de l’info-bulle sera « Cliquez pour activer/désactiver la protection de la feuille de calcul ».</span><span class="sxs-lookup"><span data-stu-id="eedea-294">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="eedea-295">Néanmoins, l’élément `resid` doit être « ProtectionButtonToolTip ».</span><span class="sxs-lookup"><span data-stu-id="eedea-295">But the `resid` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="eedea-296">Lorsque vous avez terminé, l’élément `SuperTip` doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="eedea-296">So, when you are done, the `SuperTip` element should look like this:</span></span>

    ```xml
    <Supertip>
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE]
   > <span data-ttu-id="eedea-p139">Dans un complément de production, vous n’utiliseriez pas la même icône pour deux boutons différents, mais pour simplifier ce didacticiel, nous allons le faire. Par conséquent, le balisage `Icon` de notre nouvel élément `Control` est simplement une copie de l’élément `Icon` provenant de l’élément `Control` existant.</span><span class="sxs-lookup"><span data-stu-id="eedea-p139">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span>

6. <span data-ttu-id="eedea-299">Le type de l’élément `Action` se trouvant à l’intérieur de l’élément `Control` d’origine est défini sur `ShowTaskpane`, mais notre nouveau bouton ne va pas ouvrir un volet Office, il va exécuter une fonction personnalisée que vous allez créer à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="eedea-299">The `Action` element inside the original `Control` element has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="eedea-300">Remplacez donc `TODO5` par `ExecuteFunction`, qui correspond au type d’action des boutons qui déclenchent des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="eedea-300">So, replace `TODO5` with `ExecuteFunction`, which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="eedea-301">Létiquette d’ouverture de l’élément `Action` doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="eedea-301">The opening tag for the `Action` element should look like this:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="eedea-p141">L’élément `Action` d’origine possède des éléments enfants qui spécifient un ID de volet Office ainsi qu’une URL de la page qui doit être ouverte dans le volet Office. Toutefois, un élément `Action` de type `ExecuteFunction` comporte un élément enfant unique qui nomme la fonction que le contrôle exécute. Vous créerez cette fonction à une étape ultérieure, et la nommerez `toggleProtection`. Par conséquent, remplacez `TODO6` par le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="eedea-p141">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>

    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="eedea-306">Le balisage `Control` complet doit à présent ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="eedea-306">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="eedea-307">Faites défiler vers le bas jusqu’à la section `Resources` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="eedea-307">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="eedea-308">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="eedea-308">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="eedea-309">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="eedea-309">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="eedea-310">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eedea-310">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="eedea-311">Création de la fonction qui protège la feuille</span><span class="sxs-lookup"><span data-stu-id="eedea-311">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="eedea-312">Ouvrez le fichier **.\commands\commands.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-312">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="eedea-313">Ajoutez la fonction suivante immédiatement après la fonction `action`.</span><span class="sxs-lookup"><span data-stu-id="eedea-313">Add the following function immediately after the `action` function.</span></span> <span data-ttu-id="eedea-314">Notez que nous spécifions un paramètre `args` pour la fonction et que la toute dernière ligne de la fonction appelle `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="eedea-314">Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`.</span></span> <span data-ttu-id="eedea-315">Il s’agit d’une condition requise pour toutes les commandes de type **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="eedea-315">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="eedea-316">Elle signale à l’application cliente Office que la fonction est terminée et que l’interface utilisateur est à nouveau réactive.</span><span class="sxs-lookup"><span data-stu-id="eedea-316">It signals the Office client application that the function has finished and the UI can become responsive again.</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {

            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="eedea-317">Ajoutez la ligne suivante à la fin du fichier :</span><span class="sxs-lookup"><span data-stu-id="eedea-317">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="eedea-318">À l’intérieur de la fonction `toggleProtection`, remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-318">Within the `toggleProtection` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="eedea-319">Ce code utilise la propriété de protection de l’objet de feuille de calcul dans un modèle de bouton bascule standard.</span><span class="sxs-lookup"><span data-stu-id="eedea-319">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="eedea-320">L’élément `TODO2` sera expliqué dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="eedea-320">The `TODO2` will be explained in the next section.</span></span>

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

    if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="eedea-321">Ajoutez du code pour récupérer des propriétés de document dans les objets de script du volet Office</span><span class="sxs-lookup"><span data-stu-id="eedea-321">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="eedea-322">Dans chaque fonction que vous avez créée dans ce didacticiel jusqu’à présent, vous avez mis en file d’attente les commandes pour *écrire* dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-322">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="eedea-323">Chaque fonction se terminait par un appel de la méthode `context.sync()`, qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées.</span><span class="sxs-lookup"><span data-stu-id="eedea-323">Each function ended with a call to the `context.sync()` method, which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="eedea-324">Toutefois, le code que vous avez ajouté dans la dernière étape appelle la `sheet.protection.protected property`.</span><span class="sxs-lookup"><span data-stu-id="eedea-324">However, the code you added in the last step calls the `sheet.protection.protected property`.</span></span> <span data-ttu-id="eedea-325">C’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `sheet` est uniquement un objet de proxy qui existe dans le script de votre volet Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-325">This is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="eedea-326">L’objet proxy ne connaît pas l’état réel de la protection du document. par conséquent, sa propriété `protection.protected` ne peut pas avoir de valeur réelle.</span><span class="sxs-lookup"><span data-stu-id="eedea-326">The proxy object doesn't know the actual protection state of the document, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="eedea-327">Pour éviter une erreur d’exception, vous devez d’abord récupérer l’état de protection du document et l’utiliser pour déterminer la valeur de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="eedea-327">To avoid an exception error, you must first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="eedea-328">Ce processus de récupération comporte trois étapes :</span><span class="sxs-lookup"><span data-stu-id="eedea-328">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="eedea-329">Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.</span><span class="sxs-lookup"><span data-stu-id="eedea-329">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="eedea-330">Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.</span><span class="sxs-lookup"><span data-stu-id="eedea-330">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="eedea-331">Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.</span><span class="sxs-lookup"><span data-stu-id="eedea-331">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="eedea-332">Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-332">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="eedea-333">À l’intérieur de la fonction `toggleProtection`, remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-333">Within the `toggleProtection` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="eedea-334">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-334">Note:</span></span>

   - <span data-ttu-id="eedea-p146">Chaque objet Excel possède une méthode `load`. Vous spécifiez les propriétés de l’objet que vous voulez lire dans le paramètre en tant que chaîne de noms séparés par des virgules. Dans ce cas, la propriété que vous devez lire est une sous-propriété de la propriété `protection`. Pour référence la sous-propriété, procédez presque exactement de la même façon que vous le feriez à n’importe quel autre emplacement de votre code, sauf que vous devez utiliser une barre oblique (« / ») au lieu d’un point « . ».</span><span class="sxs-lookup"><span data-stu-id="eedea-p146">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="eedea-339">Pour être sûr que la logique de bouton bascule, qui lit `sheet.protection.protected`, ne s’exécute pas tant que la synchronisation (`sync`) n’est pas terminée et que l’élément `sheet.protection.protected` n’a pas été affecté à la valeur correcte récupérée à partir du document, elle sera déplacée (à l’étape suivante) dans une fonction `then` qui ne s’exécutera pas tant que la synchronisation (`sync`) ne sera pas terminée.</span><span class="sxs-lookup"><span data-stu-id="eedea-339">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span>

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ```

2. <span data-ttu-id="eedea-p147">Il n’est pas possible que deux instructions `return` se trouvent dans le même chemin de code, donc supprimez la dernière ligne `return context.sync();` à la fin de la fonction `Excel.run`. Vous ajouterez un nouvel élément final `context.sync` dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="eedea-p147">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="eedea-342">Coupez la structurer `if ... else` dans la fonction `toggleProtection` et collez-la à la place de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="eedea-342">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="eedea-p148">Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="eedea-p148">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="eedea-345">Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que `sheet.protection.unprotect()` ou `sheet.protection.protect()` n’a pas été mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="eedea-345">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="eedea-346">La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="eedea-346">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="eedea-347">Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="eedea-347">When you are done, the entire function should look like the following:</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
          var sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

5. <span data-ttu-id="eedea-348">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-348">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-349">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-349">Test the add-in</span></span>

1. <span data-ttu-id="eedea-350">Fermez toutes les applications Office, y compris Excel.</span><span class="sxs-lookup"><span data-stu-id="eedea-350">Close all Office applications, including Excel.</span></span>

2. <span data-ttu-id="eedea-p149">Supprimez le cache Office en supprimant le contenu (tous les fichiers et sous-dossiers) du dossier cache. Cela est nécessaire pour effacer complètement l’ancienne version du complément de l’application cliente.</span><span class="sxs-lookup"><span data-stu-id="eedea-p149">Delete the Office cache by deleting the contents (all the files and subfolders) of the cache folder. This is necessary to completely clear the old version of the add-in from the client application.</span></span>

    - <span data-ttu-id="eedea-353">Pour Windows : `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="eedea-353">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="eedea-354">Pour Mac : `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="eedea-354">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

      > [!NOTE]
      > <span data-ttu-id="eedea-355">Si ce dossier n’existe pas, recherchez les dossiers suivants et, le cas échéant, supprimez le contenu du dossier :</span><span class="sxs-lookup"><span data-stu-id="eedea-355">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
      >  - <span data-ttu-id="eedea-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` où `{host}` est l’application Office (par exemple, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="eedea-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - <span data-ttu-id="eedea-357">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` où `{host}` est l’application Office (par exemple, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="eedea-357">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="eedea-358">Si le serveur web local est déjà en cours d’exécution, arrêtez-le en fermant la fenêtre de commande du nœud.</span><span class="sxs-lookup"><span data-stu-id="eedea-358">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="eedea-359">Étant donné que votre fichier manifeste a été mis à jour, vous devez à nouveau charger une version test du complément à l’aide du fichier manifeste mis à jour.</span><span class="sxs-lookup"><span data-stu-id="eedea-359">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file.</span></span> <span data-ttu-id="eedea-360">Démarrez le serveur web local et chargez indépendamment votre complément :</span><span class="sxs-lookup"><span data-stu-id="eedea-360">Start the local web server and sideload your add-in:</span></span>

    - <span data-ttu-id="eedea-361">Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-361">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="eedea-362">Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Excel avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="eedea-362">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="eedea-363">Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-363">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="eedea-364">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="eedea-364">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="eedea-365">Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="eedea-365">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="eedea-366">Sous l’onglet **Accueil** d’Excel, sélectionnez le bouton **Activer/désactiver la protection de la feuille de calcul**.</span><span class="sxs-lookup"><span data-stu-id="eedea-366">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="eedea-367">Notez que la plupart des contrôles figurant sur le ruban sont désactivés (et visuellement grisés) comme on peut le voir dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="eedea-367">Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span>

    ![Capture d’écran du ruban Excel avec le bouton de protection de la feuille de calcul activé mis en évidence.](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="eedea-p155">Choisissez une cellule comme vous le feriez si vous vouliez modifier son contenu. Excel affiche un message d'erreur indiquant que la feuille de calcul est protégée.</span><span class="sxs-lookup"><span data-stu-id="eedea-p155">Choose a cell as you would if you wanted to change its content. Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="eedea-372">Sélectionnez le bouton **Toggle Worksheet Protection** à nouveau pour réactiver les contrôles. Vous pouvez alors modifier une nouvelle fois les valeurs de cellule.</span><span class="sxs-lookup"><span data-stu-id="eedea-372">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="eedea-373">Ouvrir une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="eedea-373">Open a dialog</span></span>

<span data-ttu-id="eedea-374">Dans cette étape finale du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de boîte de dialogue au processus de volet Office et fermer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="eedea-374">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="eedea-375">Les boîtes de dialogue des compléments Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office et avec la page hôte dans le volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="eedea-375">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="eedea-376">Création de la page de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="eedea-376">Create the dialog page</span></span>

1. <span data-ttu-id="eedea-377">Dans le dossier **./src** situé à la racine du projet, créez un dossier nommé **boîtes de dialogue**.</span><span class="sxs-lookup"><span data-stu-id="eedea-377">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="eedea-378">Dans le dossier **./src/dialogs**, créez un fichier nommé **popup.html**.</span><span class="sxs-lookup"><span data-stu-id="eedea-378">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="eedea-p157">Ajoutez le balisage suivant au fichier **popup.html**. Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-p157">Add the following markup to **popup.html**. Note:</span></span>

   - <span data-ttu-id="eedea-381">La page contient un champ `<input>` dans lequel l’utilisateur doit entrer son nom et un bouton qui enverra ce nom au volet Office dans lequel il s’affiche.</span><span class="sxs-lookup"><span data-stu-id="eedea-381">The page has an `<input>` field where the user will enter their name, and a button that will send this name to the task pane where it will display.</span></span>

   - <span data-ttu-id="eedea-382">Le balisage charge un script nommé **popup.js** que vous allez créer dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="eedea-382">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="eedea-383">Il charge également la bibliothèque Office.js, car elle sera utilisée dans **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-383">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
            <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
            <input id="name-box" type="text"/><br/><br/>
            <button id="ok-button" class="ms-Button">OK</button>
        </body>
    </html>
    ```

4. <span data-ttu-id="eedea-384">Dans le dossier **./src/dialogs**, créez un fichier nommé **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-384">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="eedea-385">Ajoutez le code suivant au fichier **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-385">Add the following code to **popup.js**.</span></span> <span data-ttu-id="eedea-386">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="eedea-386">Note the following about this code:</span></span>

   - <span data-ttu-id="eedea-387">*Toutes les pages appellent les API dans la bibliothèque Office.js doivent tout d’abord vérifier que la bibliothèque est entièrement initialisée.*</span><span class="sxs-lookup"><span data-stu-id="eedea-387">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="eedea-388">La meilleure façon de procéder consiste à appeler la méthode `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="eedea-388">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="eedea-389">Si votre complément dispose de ses propres tâches d’initialisation, le code doit passer dans une méthode `then()` chaînée à l’appel de `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="eedea-389">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="eedea-390">Le code qui appelle `Office.onReady()` doit être exécuté avant tout appel à Office.js ; l’affectation se trouve donc dans un fichier de script chargé par la page, comme dans ce cas.</span><span class="sxs-lookup"><span data-stu-id="eedea-390">The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {

                // TODO1: Assign handler to the OK button.

            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="eedea-p160">Remplacez `TODO1` par le code suivant. Vous allez créer la fonction `sendStringToParentPage` lors de l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="eedea-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="eedea-393">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-393">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="eedea-394">La méthode `messageParent` transmet son paramètre à la page parent, qui est, dans ce cas, la page dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-394">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="eedea-395">Le paramètre peut être une chaîne qui inclut tous les éléments qui peuvent être sérialisés en tant que chaîne, au format XML ou JSON, ou tout type pouvant être converti en chaîne.</span><span class="sxs-lookup"><span data-stu-id="eedea-395">The parameter must be a string, which includes anything that can be serialized as a string, such as XML or JSON, or any type that can be cast to a string.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="eedea-396">Le fichier **popup.html** et le fichier **popup.js** qu’il charge s’exécutent dans un processus Microsoft Edge ou Internet Explorer 11 entièrement séparé à partir du volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="eedea-396">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="eedea-397">Si le **popup.js** était transpilé dans le même fichier **bundle.js** en tant que fichier **app.js**, le complément devrait charger deux copies du fichier **bundle.js**, ce qui irait à l’encontre de l’objectif de groupement.</span><span class="sxs-lookup"><span data-stu-id="eedea-397">If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="eedea-398">Par conséquent, ce complément ne transpile pas le fichier **popup.js** du tout.</span><span class="sxs-lookup"><span data-stu-id="eedea-398">Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="eedea-399">Mettre à jour les paramètres de configuration webapck</span><span class="sxs-lookup"><span data-stu-id="eedea-399">Update webpack config settings</span></span>

<span data-ttu-id="eedea-400">Ouvrez le fichier **webpack.config.js** situé dans le répertoire racine du projet et procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="eedea-400">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="eedea-401">Recherchez l’objet `entry` dans l’objet `config` et ajoutez une nouvelle entrée pour `popup`.</span><span class="sxs-lookup"><span data-stu-id="eedea-401">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="eedea-402">Lorsque c’est chose faite, le nouvel objet `entry` se présente comme suit :</span><span class="sxs-lookup"><span data-stu-id="eedea-402">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="eedea-403">Recherchez la matrice `plugins` dans l’objet `config` et ajoutez l’objet suivant à la fin de cette matrice.</span><span class="sxs-lookup"><span data-stu-id="eedea-403">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="eedea-404">Lorsque c’est chose faite, la nouvelle matrice `plugins` se présente comme suit :</span><span class="sxs-lookup"><span data-stu-id="eedea-404">After you've done this, the new `plugins` array will look like this:</span></span>

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    ```

3. <span data-ttu-id="eedea-405">Si le serveur web local est en cours d’exécution, arrêtez-le en fermant la fenêtre de commande du nœud.</span><span class="sxs-lookup"><span data-stu-id="eedea-405">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="eedea-406">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-406">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="eedea-407">Ouverture de la boîte de dialogue à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="eedea-407">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="eedea-408">Ouvrez le fichier **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="eedea-408">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="eedea-409">Recherchez l’élément `<button>` du bouton `freeze-header`, puis ajoutez la balise suivante après cette ligne :</span><span class="sxs-lookup"><span data-stu-id="eedea-409">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="eedea-410">La boîte de dialogue invitera l’utilisateur à saisir son nom et transmettra ce nom au volet Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-410">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="eedea-411">Le volet Office s’affichera dans une étiquette.</span><span class="sxs-lookup"><span data-stu-id="eedea-411">The task pane will display it in a label.</span></span> <span data-ttu-id="eedea-412">Immédiatement après le `button` que vous venez d’ajouter, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="eedea-412">Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="eedea-413">Ouvrez le fichier **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="eedea-413">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="eedea-414">Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `freeze-header`, puis ajoutez le code suivant après cette ligne.</span><span class="sxs-lookup"><span data-stu-id="eedea-414">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line.</span></span> <span data-ttu-id="eedea-415">Vous créerez la méthode `openDialog` lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="eedea-415">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="eedea-p165">Ajoutez la déclaration suivante à la fin du fichier. Cette variable est utilisée pour contenir un objet dans le contexte d'exécution de la page parent qui agit comme un intermédiaire vers le contexte d'exécution de la page de dialogue.</span><span class="sxs-lookup"><span data-stu-id="eedea-p165">Add the following declaration to the end of the file. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="eedea-418">Ajoutez la fonction suivante à la fin du fichier (après la déclaration de `dialog`).</span><span class="sxs-lookup"><span data-stu-id="eedea-418">Add the following function to the end of the file (after the declaration of `dialog`).</span></span> <span data-ttu-id="eedea-419">Le plus important à remarquer à propos de ce code est ce qui ne s’y trouve *pas* : il n’y a aucun appel de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="eedea-419">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="eedea-420">Cela est dû au fait que l’API d’ouverture d’une boîte de dialogue est partagée par toutes les applications Office, elle fait donc partie de l’API commune JavaScript Office, pas de l’API propre à Excel.</span><span class="sxs-lookup"><span data-stu-id="eedea-420">This is because the API to open a dialog is shared among all Office applications, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="eedea-p167">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="eedea-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="eedea-423">La méthode `displayDialogAsync` ouvre une boîte de dialogue au centre de l’écran.</span><span class="sxs-lookup"><span data-stu-id="eedea-423">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="eedea-424">Le premier paramètre est l’URL de la page à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="eedea-424">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="eedea-p168">Le deuxième paramètre transmet les options. `height` et `width` sont des pourcentages de la taille de la fenêtre de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="eedea-427">Traitement du message à partir de la boîte de dialogue et fermeture de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="eedea-427">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="eedea-428">Dans la fonction `openDialog` dans le fichier **./src/taskpane/taskpane.js**, remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="eedea-428">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code.</span></span> <span data-ttu-id="eedea-429">Remarque :</span><span class="sxs-lookup"><span data-stu-id="eedea-429">Note:</span></span>

   - <span data-ttu-id="eedea-430">Le rappel est exécuté immédiatement après que la boîte de dialogue s’est ouverte correctement et avant que l’utilisateur a pris une quelconque action dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="eedea-430">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="eedea-431">`result.value` représente l’objet qui agit comme un intermédiaire entre les contextes d’exécution des pages parent et de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="eedea-431">The `result.value` is the object that acts as an intermediary between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="eedea-p170">La fonction `processMessage` sera créée à une étape ultérieure. Ce gestionnaire traitera toutes les valeurs envoyées par la page de boîte de dialogue avec les appels de la fonction `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="eedea-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="eedea-434">Ajoutez la fonction suivante après la fonction `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="eedea-434">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="eedea-435">Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.</span><span class="sxs-lookup"><span data-stu-id="eedea-435">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="eedea-436">Test du complément</span><span class="sxs-lookup"><span data-stu-id="eedea-436">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="eedea-437">Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.</span><span class="sxs-lookup"><span data-stu-id="eedea-437">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="eedea-438">Sélectionnez le bouton **Boîte de dialogue Ouvrir** dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="eedea-438">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="eedea-p171">Lorsque la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la. Notez que vous pouvez interagir avec la feuille de calcul et appuyer sur d'autres boutons du volet des tâches, mais que vous ne pouvez pas lancer une deuxième boîte de dialogue à partir de la même page du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="eedea-p171">While the dialog is open, drag it and resize it. Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="eedea-441">Dans la boîte de dialogue, entrez un nom et appuyez sur le bouton **OK**.</span><span class="sxs-lookup"><span data-stu-id="eedea-441">In the dialog, enter a name and choose the **OK** button.</span></span> <span data-ttu-id="eedea-442">Ce nom apparaît sur le volet Office et la boîte de dialogue se ferme.</span><span class="sxs-lookup"><span data-stu-id="eedea-442">The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="eedea-p173">Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction `processMessage`. Ensuite, répétez les étapes de cette section. La boîte de dialogue reste ouverte et vous pouvez modifier le nom. Vous pouvez la fermer manuellement en appuyant sur la croix (**X**) en haut à droite.</span><span class="sxs-lookup"><span data-stu-id="eedea-p173">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Capture d’écran d’Excel, avec un bouton Ouvrir la boîte de dialogue visible dans le volet Office du complément et une boîte de dialogue affichée sur la feuille de calcul.](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="eedea-448">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="eedea-448">Next steps</span></span>

<span data-ttu-id="eedea-449">Ce didacticiel vous apprend à créer un complément Excel qui interagit avec des tableaux, des graphiques (chart), des feuilles de calcul et des boîtes de dialogue dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="eedea-449">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="eedea-450">Pour en savoir plus sur le développement des complément Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="eedea-450">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="eedea-451">Présentation des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="eedea-451">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="eedea-452">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eedea-452">See also</span></span>

- [<span data-ttu-id="eedea-453">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="eedea-453">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="eedea-454">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="eedea-454">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="eedea-455">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="eedea-455">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
