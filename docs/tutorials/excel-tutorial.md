---
title: Didacticiel sur le complément Excel
description: Dans ce didacticiel, vous allez développer un complément Excel qui crée, remplit, filtre et trie un tableau, crée un graphique, fige un en-tête de tableau, protège une feuille de calcul et ouvre une boîte de dialogue.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 3ad286e248b60afa16d4c18c9090e54e9c44cc39
ms.sourcegitcommit: f4fa1a0187466ea136009d1fe48ec67e4312c934
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408861"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a>Didacticiel : Créer un complément de volet de tâches de Excel

Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:

> [!div class="checklist"]
>
> - Crée un tableau
> - Filtres et tris un tableau
> - Crée un graphique (Chart)
> - Figer une en-tête de tableau
> - Protège une feuille de calcul
> - Ouvrir une boîte de dialogue

> [!TIP]
> Si vous avez déjà exécuté le démarrage rapide [Créer votre premier complément du volet des tâches d’Excel](../quickstarts/excel-quickstart-jquery.md) à l’aide du générateur Yeoman et que vous souhaitez utiliser ce projet comme point de départ pour ce didacticiel, accédez directement à la section [Créer un tableau](#create-a-table) pour commencer ce didacticiel.

## <a name="prerequisites"></a>Configuration requise

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a>Créer votre projet de complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project`
- **Sélectionnez un type de script :** `Javascript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

![Capture d’écran de l’interface de ligne de commande du générateur de compléments Yeoman Office](../images/yo-office-excel.png)

Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a>Créer un tableau

Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.

### <a name="code-the-add-in"></a>Codage du complément

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier **./src/taskpane/taskpane.html**.  Ce fichier contient le balisage HTML du volet des tâches.

3. Recherchez l’élément `<main>` et supprimez toutes les lignes qui apparaissent après la balise d’ouverture `<main>` et avant la balise de fermeture `</main>`.

4. Ajoutez la balise suivante juste après la balise d’ouverture `<main>` :

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. Ouvrez le fichier **./src/taskpane/taskpane.js**. Ce fichier contient le code API JavaScript Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.

6. Supprimez toutes les références au bouton `run` et à la fonction `run()` en procédant comme suit :

    - Recherchez et supprimez la ligne `document.getElementById("run").onclick = run;`.

    - Recherchez et supprimez complètement la fonction `run()`.

7. Dans l’appel de la méthode `Office.onReady`, recherchez la ligne `if (info.host === Office.HostType.Excel) {` et ajoutez le code suivant immédiatement après cette ligne. Remarque :

    - La première partie de ce code détermine si la version d’Excel de l’utilisateur prend en charge une version d’Excel.js qui inclut toutes les API que cette série de didacticiels utilisera. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur qui appellerait des API non prises en charge. Cela permettra à l’utilisateur de continuer à utiliser les parties du complément qui sont prises en charge par leur version d’Excel.

    - La deuxième partie de ce code ajoute un gestionnaire d’événements pour le bouton `create-table`.

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. Ajoutez la fonction suivante à la fin du fichier. Remarque :

    - Votre logique métier Excel.js est ajoutée à la fonction qui est transmise à `Excel.run`. Cette logique n’est pas exécutée immédiatement. En effet, elle est ajoutée à une file d’attente de commandes.

    - La méthode `context.sync` envoie toutes les commandes en file d’attente vers Excel pour exécution.

    - L’élément `Excel.run` est suivi par un bloc `catch`. Il s’agit d’une pratique recommandée que vous devez toujours suivre. 

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

9. Dans la fonction `createTable()`, remplacez `TODO1` par le code suivant. Remarque :

    - Le code crée un tableau à l’aide de la méthode `add` des collections de tableau d’une feuille de calcul, qui existe toujours même si elle est vide. Il s’agit de la méthode standard de création d’objets Excel.js. Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel. Au lieu de cela, vous l’ajoutez à un objet de la collection parent.

    - Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau. La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes. Il s’agit d’un modèle courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.

    - Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. Dans la fonction `createTable()`, remplacez `TODO2` par le code suivant. Remarque :

    - Les valeurs de cellule d’une plage sont définies avec un tableau de tableaux.

    - Les nouvelles lignes sont créées dans un tableau en appelant la méthode `add` de la collection de ligne du tableau. Vous pouvez ajouter plusieurs lignes dans un seul appel `add` en incluant plusieurs tableaux de valeurs de cellule dans le tableau parent transmis en tant que deuxième paramètre.

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

11. Dans la fonction`createTable()`, remplacez `TODO3` par le code suivant. Remarque :

    - Le code recherche une référence à la colonne **Amount** en transmettant son index de base zéro à la méthode `getItemAt` de collection de colonnes du tableau.

        > [!NOTE]
        > Les objets de collection Excel.js, tels que `TableCollection`, `WorksheetCollection` et `TableColumnCollection` ont une propriété `items` qui correspond à un tableau de types d’objet enfant, comme `Table` ou `Worksheet` ou `TableColumn` ; mais un objet `*Collection` n’est pas lui-même un tableau.

    - Le code définit ensuite la plage de la colonne **Amount** sous la forme Euros à la deuxième décimale.

    - Enfin, il s’assure que la largeur des colonnes et la hauteur des lignes sont assez grandes pour contenir l’élément de données le plus long (ou le plus haut). Notez que le code doit rechercher des objets `Range` à mettre en forme. Les objets `TableColumn` et `TableRow` n’ont pas de propriétés de mise en forme.

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

    > [!NOTE]
    > Les compléments Office doivent utiliser HTTPS, pas HTTP, même lorsque vous effectuez le développement. Si vous êtes invité à installer un certificat après avoir exécuté l’une des commandes suivantes, acceptez l’invite pour installer le certificat fourni par le générateur Yeoman.

    > [!TIP]
    > Si vous testez votre complément sur Mac, exécutez la commande suivante dans le répertoire racine de votre projet avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Cela démarre le serveur web local (s’il n’est pas déjà en cours d’exécution) et ouvre Excel avec votre complément chargé.

        ```command&nbsp;line
        npm start
        ```

    - Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

        ```command&nbsp;line
        npm run start:web
        ```

        Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran du menu Accueil d’Excel avec le bouton Afficher le volet Office mis en évidence](../images/excel-quickstart-addin-3b.png)

3. Dans le volet Office, sélectionnez le bouton **Créer un tableau**.

    ![Capture d’écran d’Excel, affichant le volet Office Complément avec le bouton créer un tableau, et un tableau dans la feuille de calcul rempli avec les données de date, de commerçant, de catégorie et de montant](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a>Filtrer et trier un tableau

Dans cette étape du didacticiel, vous allez filtrer et trier le tableau que vous avez créé précédemment.

### <a name="filter-the-table"></a>Filtrage du tableau

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

2. Recherchez l’élément `<button>` du bouton `create-table`, puis ajoutez la balise suivante après cette ligne :

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. Ouvrez le fichier **./src/taskpane/taskpane.js**.

4. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `create-table`, puis ajoutez le code suivant après cette ligne :

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. Ajoutez la fonction suivante à la fin du fichier :

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

6. Dans la fonction `filterTable()`, remplacez `TODO1` par le code suivant. Remarque :

   - Le code obtient tout d’abord une référence à la colonne à filtrer en transférant le nom de la colonne à la méthode `getItem`, au lieu de transmettre son index à la méthode `getItemAt` comme le fait la méthode `createTable`. Puisque les utilisateurs peuvent déplacer des colonnes de tableau, la colonne d’un index donné peut être modifiée après la création du tableau. Par conséquent, il est préférable d’utiliser le nom de la colonne pour obtenir une référence de la colonne. Dans le didacticiel précédent, nous avons utilisé la méthode `getItemAt` en toute sécurité, car nous l’avons utilisée dans la même méthode que celle qui crée le tableau, il n’y a donc aucune chance qu’un utilisateur ait déplacé la colonne.

   - La méthode `applyValuesFilter` est l’une des nombreuses méthodes de filtrage sur l’objet `Filter`.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ```

### <a name="sort-the-table"></a>Tri du tableau

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

2. Recherchez l’élément `<button>` du bouton `filter-table`, puis ajoutez la balise suivante après cette ligne :

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. Ouvrez le fichier **./src/taskpane/taskpane.js**.

4. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `filter-table`, puis ajoutez le code suivant après cette ligne :

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. Ajoutez la fonction suivante à la fin du fichier :

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

6. Dans la fonction `sortTable()`, remplacez `TODO1` par le code suivant. Remarque :

   - Le code crée un tableau d’objets `SortField` qui ne comporte qu’un seul membre puisque le complément ne trie que la colonne Marchand.

   - La propriété `key` d’un objet `SortField` est l’index de base zéro de la colonne utilisée pour le tri. Les lignes du tableau sont triées en fonction des valeurs de la colonne référencée.

   - Le membre `sort` d’un `Table` est un objet `TableSort`, et non une méthode. Les `SortField` sont transmis à la méthode `apply` de l’objet `TableSort`.

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

7. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

3. Si la tableau que vous avez ajoutée précédemment dans ce didacticiel ne figure pas dans la feuille de calcul ouverte, sélectionnez le bouton **Créer un tableau** dans le volet Office.

4. Choisissez le bouton **Filtrer le tableau** et le bouton **Trier le tableau** dans n’importe quel ordre.

    ![Capture d’écran d’Excel, avec les boutons Filtrer le tableau et Trier le tableau mis en évidence dans le volet Office Complément](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a>Création d’un graphique (chart)

Dans cette étape du didacticiel, vous créerez un graphique à l’aide de données provenant du tableau précédemment créé, puis vous mettrez en forme le graphique.

### <a name="chart-a-chart-using-table-data"></a>Un graphique à l’aide de données du tableau de graphique (chart)

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

2. Recherchez l’élément `<button>` du bouton `sort-table`, puis ajoutez la balise suivante après cette ligne : 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. Ouvrez le fichier **./src/taskpane/taskpane.js**.

4. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `sort-table`, puis ajoutez le code suivant après cette ligne :

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. Ajoutez la fonction suivante à la fin du fichier :

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

6. Dans la fonction `createChart()`, remplacez `TODO1` par le code suivant. Notez que pour exclure la ligne d’en-tête, le code utilise la méthode `Table.getDataBodyRange` pour obtenir la plage de données que vous souhaitez représenter sous forme graphique au lieu de la méthode `getRange`.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. Dans la fonction `createChart()`, remplacez `TODO2` par le code suivant. Notez les paramètres suivants :

   - Le premier paramètre transmis à la méthode `add` spécifie le type de graphique. Il en existe plusieurs dizaines de types.

   - Le deuxième paramètre spécifie la plage de données à inclure dans le graphique.

   - Le troisième paramètre détermine si une série de points de données de la table doit être représentée par ligne ou par colonne. L’option `auto` indique à Excel de décider de la meilleure méthode.

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. Dans la fonction `createChart()`, remplacez `TODO3` par le code suivant. La majorité de ce code est explicite. Remarque :

   - Les paramètres de la méthode `setPosition` spécifient les cellules situées en haut à gauche et en bas à droite de la zone de feuille de calcul devant contenir le graphique. Excel peut ajuster des éléments, tels que la largeur de ligne pour que le graphique s’affiche correctement dans l’espace attribué.

   - Une « série » est un ensemble de points de données dans une colonne du tableau. Étant donné qu’il n’existe qu’une seule colonne autre que de type chaîne dans le tableau, Excel déduit que la colonne est la seule colonne de points de données pour le graphique. Il interprète les autres colonnes comme des étiquettes de graphique. Par conséquent, il y aura simplement une série dans le graphique et un index 0. Il s’agit de celle à étiqueter avec « Valeur en &euro; ».

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in &euro;';
    ```

9. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

3. Si la tableau que vous avez ajoutée précédemment dans ce didacticiel ne figure pas dans la feuille de calcul ouverte, sélectionnez le bouton **Créer un tableau**, puis le bouton **Filtrer un tableau** et le bouton **Trier un tableau** dans n’importe quel ordre..

4. Sélectionnez le bouton **Créer un graphique**. Un graphique est créé dans lequel seules les données provenant des lignes filtrées sont incluses. Les étiquettes sur les points de données en bas sont organisées selon l’ordre de tri du graphique, à savoir les noms de marchand par ordre alphabétique inversé.

    ![Capture d’écran d’Excel avec un bouton créer un graphique visible dans le volet Office Complément, et un graphique dans la feuille de calcul affichant les données de dépenses de courses et d’éducation](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a>Figer un en-tête de tableau

Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée. Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas.

### <a name="freeze-the-tables-header-row"></a>Figer la ligne d’en-tête du tableau

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

2. Recherchez l’élément `<button>` du bouton `create-chart`, puis ajoutez la balise suivante après cette ligne :

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. Ouvrez le fichier **./src/taskpane/taskpane.js**.

4. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `create-chart`, puis ajoutez le code suivant après cette ligne :

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. Ajoutez la fonction suivante à la fin du fichier :

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

6. Dans la fonction`freezeHeader()`, remplacez `TODO1` par le code suivant. Remarque :

   - La collection `Worksheet.freezePanes` est un ensemble de volets de la feuille de calcul qui sont épinglés, c’est-à-dire figés, lorsque vous faites défiler la feuille de calcul.

   - La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être épinglées. `1` est transmis pour épingler la première rangée.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

3. Si le tableau que vous avez ajouté précédemment dans ce didacticiel est présent dans la feuille de calcul, supprimez-le.

4. Dans le volet Office, sélectionnez le bouton **Créer un tableau**.

5. Dans le volet Office, sélectionnez le bouton **Figer l’en-tête**.

6. Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.

    ![Capture d’écran illustrant une feuille de calcul Excel avec un en-tête de tableau figé](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a>Protéger une feuille de calcul

Au cours de cette étape, vous allez ajouter un bouton au ruban pour activer ou désactiver la protection de la feuille de calcul.

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>Configuration du manifeste pour ajouter un deuxième bouton de ruban

1. Ouvrez le fichier manifeste **./manifest.xml**.

2. Recherchez l’élément `<Control>`. Cet élément définit le bouton **Afficher le volet des tâches** sur le ruban **Accueil** que vous avez utilisé pour lancer le complément. Nous allons ajouter un deuxième bouton au même groupe sur le ruban **Accueil**. Entre la balise de fermeture `</Control>` et la balise de fermeture `</Group>`, ajoutez le balisage suivant.

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

3. Dans le XML que vous venez d’ajouter au fichier manifeste, remplacez `TODO1` par une chaîne qui donne au bouton un ID unique dans ce fichier manifeste. Puisque notre bouton va activer et désactiver la protection de la feuille de calcul, utilisez « ToggleProtection ». Lorsque vous avez terminé, la balise d’ouverture de l’élément `Control` devrait ressembler à ceci :

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. Les trois prochains `TODO` définissent les ID de ressource, ou les `resid`. Une ressource est une chaîne, et vous créerez ces trois chaînes dans une étape ultérieure. Pour l’instant, vous devez donner des identifiants aux ressources. Le libellé du bouton doit indiquer « Toggle Protection », mais l’*ID* de cette chaîne doit être « ProtectionButtonLabel », donc l’élément`Label` doit ressembler à ceci :

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. L’élément `SuperTip` définit l’info-bulle du bouton. Le titre de l’info-bulle doit être le même que celui du bouton, nous utilisons donc le même ID de ressource : « ProtectionButtonLabel ». La description de l’info-bulle sera « Cliquez pour activer ou désactiver la protection de la feuille de calcul ». Mais le `resid` doit être « ProtectionButtonToolTip ». Ainsi, lorsque vous avez terminé, l’élément `SuperTip` devrait ressembler à ceci :

    ```xml
    <Supertip>
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE]
   > Dans un complément de production, vous n’utilisez pas la même icône pour deux boutons différents, mais pour simplifier ce didacticiel, nous allons le faire. Par conséquent, le balisage `Icon` de notre nouvel élément `Control` est simplement une copie de l’élément `Icon` provenant de l’élément `Control` existant.

6. L’élément `Action` à l’intérieur de l’élément `Control` d’origine a son type défini sur `ShowTaskpane`, mais notre nouveau bouton n’ouvrira pas le volet des tâches. Il va exécuter une fonction personnalisée que vous créez lors d’une étape ultérieure. Remplacez donc `TODO5` par `ExecuteFunction`, qui est le type d’action des boutons qui déclenchent des fonctions personnalisées. La balise d’ouverture de l’élément `Action` doit ressembler à ceci :

    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. L’élément `Action` d’origine possède des éléments enfants qui spécifient un ID de volet des tâches ainsi qu’une URL de la page qui doit être ouverte dans le volet des tâches. Toutefois, un élément `Action` de type `ExecuteFunction` comporte un élément enfant unique qui nomme la fonction que le contrôle exécute. Vous créerez cette fonction à une étape ultérieure, et la nommerez `toggleProtection`. Par conséquent, remplacez `TODO6` par le balisage suivant :

    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    Le balisage `Control` complet doit à présent ressembler à ce qui suit :

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

8. Faites défiler vers le bas jusqu’à la section `Resources` du manifeste.

9. Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:ShortStrings`.

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:LongStrings`.

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. Enregistrez le fichier.

### <a name="create-the-function-that-protects-the-sheet"></a>Création de la fonction qui protège la feuille

1. Ouvrez le fichier **.\commands\commands.js**.

2. Ajoutez la fonction suivante immédiatement après la fonction `action`. Notez que nous spécifions un paramètre `args` pour la fonction et la toute dernière ligne des appels de fonction `args.completed`. Il s’agit d’une exigence pour toutes les commandes complémentaires de type **ExecuteFunction**. Il signale à l’application cliente Office que la fonction est terminée et que l’interface utilisateur peut redevenir réactive.

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

3. Ajoutez la ligne suivante à la fin du fichier :

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. Dans la fonction `toggleProtection`, remplacez `TODO1` par le code suivant. Ce code utilise la propriété de protection de l’objet de feuille de calcul dans un modèle de bascule standard. La tâche `TODO2` sera expliquée dans la section suivante.

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Ajouter du code pour récupérer des propriétés de document dans les objets de script du volet des tâches

Dans chaque fonction que vous avez créée dans ce didacticiel jusqu’à présent, vous avez mis des commandes en file d’attente pour *écrire* dans le document Office. Chaque fonction s’est terminée par un appel à la méthode `context.sync()`, qui envoie les commandes en file d’attente au document à exécuter. Toutefois, le code que vous avez ajouté à la dernière étape appelle la `sheet.protection.protected property`. Il s’agit d’une différence significative par rapport aux fonctions précédentes que vous avez écrites, car l’objet `sheet` est uniquement un objet proxy qui existe dans le script de votre volet des tâches. L’objet proxy ne connaît pas l’état de protection réel du document, donc sa propriété `protection.protected` ne peut pas avoir une valeur réelle. Pour éviter une erreur d’exception, vous devez d’abord extraire l’état de protection du document et l’utiliser pour définir la valeur de `sheet.protection.protected`. Ce processus de récupération comporte trois étapes :

   1. Mettez en file d’attente une commande de chargement (c’est-à-dire, récupérer) des propriétés que votre code doit lire.

   2. Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.

   3. Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.

Ces étapes doivent être effectuées à chaque fois que votre code doit *lire* des informations provenant du document Office.

1. Dans la fonction `toggleProtection`, remplacez `TODO2` par le code suivant. remarque :

   - Chaque objet Excel possède une méthode `load`. Vous spécifiez les propriétés de l’objet que vous voulez lire dans le paramètre en tant que chaîne de noms séparés par des virgules. Dans ce cas, la propriété que vous devez lire est une sous-propriété de la propriété `protection`. Pour référencer la sous-propriété, procédez presque exactement de la même façon que vous le feriez à n’importe quel autre emplacement de votre code, mais utilisez une barre oblique (« / ») au lieu d’un point « . ».

   - Pour être sûr que la logique de bouton bascule, qui lit `sheet.protection.protected`, ne s’exécute pas tant que la synchronisation (`sync`) n’est pas terminée et que l’élément `sheet.protection.protected` n’a pas été affecté à la valeur correcte récupérée à partir du document, elle sera déplacée (à l’étape suivante) dans une fonction `then` qui ne s’exécutera pas tant que la synchronisation (`sync`) ne sera pas terminée.

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

2. Il n’est pas possible que deux instructions `return` se trouvent dans le même chemin de code, donc supprimez la dernière ligne `return context.sync();` à la fin de la fonction `Excel.run`. Vous ajouterez un nouvel élément final `context.sync` dans une étape ultérieure.

3. Coupez la structurer `if ... else` dans la fonction `toggleProtection` et collez-la à la place de `TODO3`.

4. Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :

   - Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que `sheet.protection.unprotect()` ou `sheet.protection.protect()` n’a pas été mis en file d’attente.

   - La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de `context.sync`.

    ```js
    .then(context.sync);
    ```

   Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :

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

5. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. Fermez toutes les applications Office, y compris Excel.

2. Supprimez le cache Office en supprimant le contenu (tous les fichiers et sous-dossiers) du dossier cache. Cela est nécessaire pour effacer complètement l’ancienne version du complément de l’application cliente.

    - Pour Windows : `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

    - Pour Mac : `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

      > [!NOTE]
      > Si ce dossier n’existe pas, recherchez les dossiers suivants et, le cas échéant, supprimez le contenu du dossier :
      >  - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` où `{host}` est l’application Office (par exemple, `Excel`)
      >  - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` où `{host}` est l’application Office (par exemple, `Excel`)
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`

3. Si le serveur web local est déjà en cours d’exécution, arrêtez-le en fermant la fenêtre de commande node.

4. Votre fichier manifeste étant été mis à jour, vous devez à nouveau charger votre complément à l’aide du fichier manifeste. Démarrez le serveur web local et chargez votre complément :

    - Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Cela démarre le serveur web local (s’il n’est pas déjà en cours d’exécution) et ouvre Excel avec votre complément chargé.

        ```command&nbsp;line
        npm start
        ```

    - Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

        ```command&nbsp;line
        npm run start:web
        ```

        Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

5. Sous l’onglet **Accueil** dans Excel, choisissez le bouton **Activer la protection de la feuille de calcul**. Notez que la plupart des contrôles sur le ruban sont désactivés (et visuellement grisés) comme le montre la capture d’écran suivante.

    ![Capture d’écran du ruban Excel avec le bouton Activer la protection de la feuille de calcul mis en évidence et activé. La plupart des autres boutons apparaissent en gris et sont désactivés.](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. Choisissez une cellule comme si vous souhaitiez changer son contenu. Excel affiche un message d’erreur indiquant que la feuille de calcul est protégée.

7. Sélectionnez le bouton **Activer la protection de la feuille de calcul** à nouveau pour réactiver les contrôles. Vous pouvez alors modifier une nouvelle fois les valeurs de cellule.

## <a name="open-a-dialog"></a>Ouvrir une boîte de dialogue

Dans cette dernière étape du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de dialogue au processus du volet des tâches et fermer la boîte de dialogue. Les boîtes de dialogue du complément Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office et avec la page hôte dans le volet des tâches.

### <a name="create-the-dialog-page"></a>Créer la page de boîte de dialogue

1. Dans le dossier **./src** situé à la racine du projet, créez un dossier nommé **boîtes de dialogue**.

2. Dans le dossier **./src/dialogs**, créez un fichier nommé **popup.html**.

3. Ajoutez le balisage suivant au fichier **popup.html**. Remarque :

   - La page contient un champ `<input>` dans lequel l’utilisateur doit entrer son nom. Un bouton enverra ce nom au volet des tâches dans lequel il s’affiche.

   - Le balisage charge un script nommé **popup.js** que vous allez créer dans une étape ultérieure.

   - Il charge également la bibliothèque Office.js, car elle sera utilisée dans **popup.js**.

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
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

4. Dans le dossier **./src/dialogs**, créez un fichier nommé **popup.js**.

5. Ajoutez le code suivant au fichier **popup.js**. Notez ce qui suit à propos de ce code :

   - *Chaque page qui appelle des API dans la bibliothèque Office.js doit d’abord s’assurer que la bibliothèque est entièrement initialisée.* La meilleure façon de procéder est d’appeler la méthode `Office.onReady()`. Si votre complément a ses propres tâches d’initialisation, le code doit être envoyé dans une méthode `then()` qui est enchaînée à l’appel `Office.onReady()`. L’appel de `Office.onReady()` doit s’exécuter avant tout appel à Office.js. Par conséquent, l’affectation est dans un fichier de script qui est chargé par la page, comme c’est le cas ici.

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

6. Remplacez `TODO1` par le code suivant. Vous allez créer la fonction `sendStringToParentPage` lors de l’étape suivante.

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. Remplacez `TODO2` par le code suivant. La méthode `messageParent` transmet son paramètre à la page parent, qui est, dans ce cas, la page dans le volet des tâches. Le paramètre peut être une valeur booléenne ou une chaîne qui inclut tous les éléments qui peuvent être sérialisés en tant que chaîne, au format XML ou JSON.

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> Le fichier **popup.html** et le fichier **popup.js** qu’il charge, s’exécutent dans un processus Microsoft Edge ou Internet Explorer 11 entièrement distinct du volet des tâches du complément. Si le fichier **popup.js** était transmis dans le même fichier **bundle.js** que le fichier **app.js**, alors le complément devrait charger deux copies du fichier **bundle.js**, ce qui va à l’encontre de l’objectif du regroupement. Par conséquent, ce complément ne transmet pas du tout le fichier **popup.js**.

### <a name="update-webpack-config-settings"></a>Mettre à jour les paramètres de configuration webapck

Ouvrez le fichier **webpack.config.js** situé dans le répertoire racine du projet et procédez comme suit.

1. Recherchez l’objet `entry` dans l’objet `config` et ajoutez une nouvelle entrée pour `popup`.

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    Lorsque c’est chose faite, le nouvel objet `entry` se présente comme suit :

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. Recherchez la matrice `plugins` dans l’objet `config` et ajoutez l’objet suivant à la fin de cette matrice.

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    Lorsque c’est chose faite, la nouvelle matrice `plugins` se présente comme suit :

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

3. Si le serveur web local est en cours d’exécution, arrêtez-le en fermant la fenêtre de commande du nœud.

4. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a>Ouverture de la boîte de dialogue à partir du volet Office

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

2. Recherchez l’élément `<button>` du bouton `freeze-header`, puis ajoutez la balise suivante après cette ligne :

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. La boîte de dialogue invitera l’utilisateur à entrer un nom et à transmettre le nom de l’utilisateur au volet Office. Le volet Office l’affichera dans une étiquette. Immédiatement après le `button` que vous venez d’ajouter, ajoutez le balisage suivant :

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. Ouvrez le fichier **./src/taskpane/taskpane.js**.

5. Dans l’appel de la méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `freeze-header` et ajoutez le code suivant après cette ligne. Vous allez créer la méthode `openDialog` lors d’une étape ultérieure.

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. Ajoutez la déclaration suivante à la fin du fichier. Cette variable est utilisée pour contenir un objet dans le contexte d’exécution de la page parente qui agit comme intermédiaire dans le contexte d’exécution de la page de dialogue.

    ```js
    var dialog = null;
    ```

7. Ajoutez la fonction suivante à la fin du fichier (après la déclaration de `dialog`). La chose importante à noter à propos de ce code est ce qui n’y est *pas* : il n’y a pas d’appel `Excel.run`. En effet, l’API permettant d’ouvrir une boîte de dialogue est partagée entre toutes les applications Office, elle fait donc partie de l’API commune JavaScript d’Office, et non de l’API spécifique à Excel.

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. Remplacez `TODO1` par le code suivant. Remarque :

   - La méthode `displayDialogAsync` ouvre une boîte de dialogue au centre de l’écran.

   - Le premier paramètre est l’URL de la page à ouvrir.

   - Le deuxième paramètre transmet les options. `height` et `width` sont des pourcentages de la taille de la fenêtre de l’application Office.

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Traiter le message à partir de la boîte de dialogue et fermer la boîte de dialogue

1. Dans la fonction `openDialog` du fichier **./src/taskpane/taskpane.js**, remplacez `TODO2` par le code suivant. Remarque :

   - Le rappel est exécuté immédiatement après que l’ouverture correcte de la boîte de dialogue et avant que l’utilisateur entreprenne une action dans la boîte de dialogue.

   - `result.value` représente l’objet qui agit comme un intermédiaire entre les contextes d’exécution des pages parent et de boîte de dialogue.

   - La fonction `processMessage` sera créée à une étape ultérieure. Ce gestionnaire traitera toutes les valeurs envoyées par la page de boîte de dialogue avec les appels de la fonction `messageParent`.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. Ajoutez la fonction suivante après la fonction `openDialog`.

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. Si le volet des tâches du complément n’est pas déjà ouvert dans Excel, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

3. Sélectionnez le bouton **Ouvrir la boîte de dialogue** dans le volet des tâches.

4. Pendant que la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la. Notez que vous pouvez interagir avec la feuille de calcul et appuyer sur d’autres boutons du volet des tâches mais vous ne pouvez pas ouvrir une deuxième boîte de dialogue à partir de la même page du volet des tâches.

5. Dans la boîte de dialogue, entrez un nom et sélectionnez le bouton **OK**. Le nom apparaît dans le volet des tâches et la boîte de dialogue se ferme.

6. Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction`processMessage`. Ensuite, répétez les étapes de cette section. La boîte de dialogue reste ouverte et vous pouvez modifier le nom. Vous pouvez la fermer manuellement en appuyant sur le bouton (**X**) en haut à droite.

    ![Capture d’écran d’Excel, avec un bouton de boîte de dialogue Ouvrir visible dans le volet Office Complément, et une boîte de dialogue affichée sur la feuille de calcul](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a>Étapes suivantes

Dans ce didacticiel, vous avez créé un complément de volet des tâches Excel qui interagit avec les tableaux, les graphiques, les feuilles de calcul et les boîtes de dialogue des classeurs Excel. Si vous souhaitez en savoir plus sur la création de compléments Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Présentation des compléments Excel](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](../excel/excel-add-ins-core-concepts.md)
