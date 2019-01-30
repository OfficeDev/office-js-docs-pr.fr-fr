---
title: Didacticiel sur le complément Excel
description: Dans ce didacticiel, vous allez développer un complément Excel qui crée, remplit, filtre et trie un tableau, crée un graphique, fige un en-tête de tableau, protège une feuille de calcul et ouvre une boîte de dialogue.
ms.date: 01/28/2019
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 6fe72a9170862dbb0c422db7d8efd3f187bf45ae
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635964"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a>Didacticiel : Créer un complément de volet de tâches de Excel

Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:

> [!div class="checklist"]
> * Crée un tableau
> * Filtres et tris un tableau
> * Crée un graphique (Chart)
> * Figer une en-tête de tableau
> * Protège une feuille de calcul
> * Ouvrir une boîte de dialogue

## <a name="prerequisites"></a>Conditions requises

Pour utiliser ce didacticiel, les logiciels suivants doivent être installés. 

- Excel 2016, version 1711 (Démarrer en un clic version 8730.1000) ou version ultérieure. Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).

- [Node](https://nodejs.org/en/) 

- [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

- Vous devez disposer d’une connexion Internet pour tester le complément dans ce didacticiel.

## <a name="create-your-add-in-project"></a>Créer votre projet de complément

Procédez comme suit pour créer le projet de complément Excel que vous souhaitez utiliser comme base pour ce didacticiel.

1. Clonez le référentiel GitHub du [didacticiel sur les compléments Excel](https://github.com/OfficeDev/Excel-Add-in-Tutorial).

2. Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

3. Exécutez la commande `npm install` pour installer les outils et les bibliothèques répertoriées dans le fichier package.json. 

4. Effectuez les étapes décrites dans la rubrique relative à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.

## <a name="create-a-table"></a>Créer un tableau

Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.

### <a name="code-the-add-in"></a>Codage du complément

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. Remplacez `TODO1` par le codage suivant :

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. Ouvrez le fichier app.js.

5. Remplacez `TODO1` par le code suivant. Ce code détermine si la version Excel de l’utilisateur prend en charge une version d’Excel.js qui inclut toutes les API utilisées dans cette série de didacticiels. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge. Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par leur version d’Excel.

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. Remplacez `TODO2` par le code suivant :

    ```js
    $('#create-table').click(createTable);
    ```

7. Remplacez `TODO3` par le code suivant. Remarque :

   - Votre logique métier Excel.js est ajoutée à la fonction qui est transmise à `Excel.run`. Cette logique n’est pas exécutée immédiatement. Au lieu de cela, elle est ajoutée à une file d’attente de commandes.

   - La méthode `context.sync` envoie toutes les commandes en file d’attente vers Excel pour exécution.

   - L’élément `Excel.run` est suivi par un bloc `catch`. Il s’agit d’une meilleure pratique que vous devez toujours suivre. 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. Remplacez `TODO4` par le code suivant. Tenez compte des informations suivantes :

   - Le code crée un tableau à l’aide de la méthode `add` de collection de tableau d’une feuille de calcul, qui existe toujours même si elle est vide. Il s’agit de la méthode standard de création d’objets Excel.js. Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel. Au lieu de cela, vous l’ajoutez à un objet de la collection parent.

   - Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau. La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes. Il s’agit d’un modèle plus courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.

   - Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. Remplacez `TODO5` par le code suivant. Remarque :

   - Les valeurs de cellule d’une plage sont définies avec un tableau de tableaux.

   - Les nouvelles lignes sont créées dans un tableau en appelant la méthode `add` de collection de ligne du tableau. Vous pouvez ajouter plusieurs lignes dans un seul appel de `add` en incluant plusieurs tableaux de valeurs de cellule dans le tableau parent transmis en tant que deuxième paramètre.

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

10. Remplacez `TODO6` par le code suivant. Remarque :

   - Le code recherche une référence à la colonne **Amount** en transmettant son index de base zéro à la méthode `getItemAt` de collection de colonnes du tableau.

     > [!NOTE]
     > Les objets de collection Excel.js, tels que `TableCollection`, `WorksheetCollection` et `TableColumnCollection` ont une propriété `items` qui correspond à un tableau de types d’objet enfant, comme `Table` ou `Worksheet` ou `TableColumn` ; mais un objet `*Collection` n’est pas lui-même un tableau.

   - Le code définit ensuite la plage de la colonne **Amount** sous la forme Euros à la deuxième décimale. 

   - Enfin, il s’assure que la largeur des colonnes et la hauteur des lignes sont assez grandes pour contenir l’élément de données le plus long (ou le plus haut). Notez que le code doit rechercher des objets `Range` à mettre en forme. Les objets `TableColumn` et `TableRow` n’ont pas de propriétés de mise en forme.

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a>Test du complément

1. Ouvrez une fenêtre Git Bash ou une invite système activée par Node.JS, et accédez au dossier **Démarrer** du projet.

2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 sur une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisée en arrière-plan par Excel pour exécuter les compléments Excel).

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Chargez une version test du complément en utilisant l’une des méthodes suivantes :

    - Windows : [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)

    - iPad et Mac : [Chargement de versions test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

5. Dans le menu **Accueil**, sélectionnez **Afficher le volet Office**.

6. Dans le volet Office, sélectionnez **Créer un tableau**.

    ![Didacticiel Excel -Créer un tableau](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a>Filtrer et trier un tableau

Dans cette étape du didacticiel, vous allez filtrer et trier le tableau que vous avez créé précédemment.

### <a name="filter-the-table"></a>Filtrage du tableau

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. Juste en dessous de la balise `div` qui contient le bouton `create-table`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-table`, ajoutez le code suivant :

    ```js
    $('#filter-table').click(filterTable);
    ```

6. Ajoutez la fonction suivante juste après la fonction `createTable`.

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

7. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - Le code obtient tout d’abord une référence à la colonne à filtrer en transférant le nom de la colonne à la méthode `getItem`, au lieu de transmettre son index à la méthode `getItemAt` comme le fait la méthode `createTable`. Puisque les utilisateurs peuvent déplacer des colonnes de tableau, la colonne d’un index donné peut être modifiée après la création du tableau. Par conséquent, il est préférable d’utiliser le nom de la colonne pour obtenir une référence de la colonne. Dans le didacticiel précédent, nous avons utilisé la méthode `getItemAt` en toute sécurité, car nous l’avons utilisée dans la même méthode que celle qui crée le tableau, il n’y a donc aucune chance qu’un utilisateur ait déplacé la colonne.

   - La méthode `applyValuesFilter` est l’une des nombreuses méthodes de filtrage sur l’objet `Filter`.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a>Tri du tableau

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `filter-table`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. En dessous de la ligne qui attribue un gestionnaire de clic au bouton `filter-table`, ajoutez le code suivant :

    ```js
    $('#sort-table').click(sortTable);
    ```

5. Ajoutez la fonction suivante après la fonction `filterTable`.

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

6. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - Le code crée un tableau d’objets `SortField` qui ne comporte qu’un seul membre puisque le complément ne trie que la colonne Merchant.

   - La propriété `key` d’un objet `SortField` est l’index de la colonne à trier qui part de zéro.

   - Le membre `sort` d’un objet `Table` est un objet `TableSort`, et non une méthode. Les objets `SortField` sont transmis à la méthode `apply` de l’objet `TableSort`.

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

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur **Ctrl+C **à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Si, pour une raison quelconque, le tableau ne se trouve pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**.

6. Choisissez les boutons **Filtrer le tableau** et **Trier le tableau** dans n’importe quel ordre.

    ![Didacticiel Excel- Filtrer et trier un tableau](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a>Création d’un graphique (chart)

Dans cette étape du didacticiel, vous créerez un graphique à l’aide de données provenant du tableau précédemment créé, puis vous mettrez en forme le graphique.

### <a name="chart-a-chart-using-table-data"></a>Un graphique à l’aide de données du tableau de graphique (chart)

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. En dessous de la balise `div` qui contient le bouton `sort-table`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. En dessous de la ligne qui attribue un gestionnaire de clic au bouton `sort-chart`, ajoutez le code suivant :

    ```js
    $('#create-chart').click(createChart);
    ```

6. Sous la fonction `sortTable`, ajoutez la fonction suivante.

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

7. Remplacez `TODO1` par le code suivant. Pour exclure la ligne d’en-tête, le code utilise la méthode `Table.getDataBodyRange` pour obtenir la plage de données à représenter sous forme de graphique à la place de la méthode `getRange`.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. Remplacez `TODO2` par le code suivant. Notez les paramètres suivants :

   - Le premier paramètre transmis à la méthode `add` spécifie le type de graphique. Il en existe plusieurs dizaines de types.

   - Le deuxième paramètre spécifie la plage de données à inclure dans le graphique.

   - Le troisième paramètre détermine si une série de points de données provenant du tableau doit être représentée sous forme de graphique par ligne ou par colonne. L’option `auto` demande à Excel de déterminer la meilleure méthode.

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. Remplacez `TODO3` par le code suivant. La majeure partie du code est explicite. Remarque :
   
   - Les paramètres de la méthode `setPosition` spécifient les cellules situées en haut à gauche et en bas à droite de la zone de feuille de calcul devant contenir le graphique. Excel peut ajuster des éléments, tels que la largeur de ligne pour que le graphique s’affiche correctement dans l’espace attribué.
   
   - Une « série » est un ensemble de points de données dans une colonne du tableau. Étant donné qu’il n’existe qu’une seule colonne autre que de type chaîne dans le tableau, Excel déduit que la colonne est la seule colonne de points de données pour le graphique. Il interprète les autres colonnes comme des étiquettes de graphique. Par conséquent, il y aura simplement une série dans le graphique et un index 0. Il s’agit de celle à étiqueter avec « Valeur en € ».

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur **Ctrl+C **à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Si pour une raison quelconque le tableau n’est pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**, puis les boutons **Filtrer le tableau** et **Trier le tableau** dans n’importe quel ordre.

6. Sélectionnez le bouton **Créer un graphique**. Un graphique est créé dans lequel seules les données provenant des lignes filtrées sont incluses. Les étiquettes sur les points de données en bas sont organisées selon l’ordre de tri du graphique, à savoir les noms de marchand par ordre alphabétique inversé.

    ![Didacticiel Excel -Créer un graphique (chart)](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a>Figer un en-tête de tableau

Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée. Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas.

### <a name="freeze-the-tables-header-row"></a>Figer la ligne d’en-tête du tableau

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. En dessous de la balise `div` qui contient le bouton `create-chart`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. En dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-chart`, ajoutez le code suivant :

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. En dessous de la fonction `createChart`, ajoutez la fonction suivante :

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

7. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - La collection `Worksheet.freezePanes` est un ensemble de volets de la feuille de calcul qui sont épinglés, c’est-à-dire figés, lorsque vous faites défiler la feuille de calcul.

   - La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être figées. Nous transmettons `1` pour figer la première ligne.

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur **Ctrl+C **à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet Office** pour rouvrir le complément.

5. Si le tableau est dans la feuille de calcul, supprimez-le.

6. Dans le volet Office, sélectionnez **Créer un tableau**.

7. Sélectionnez le bouton **Freeze Header**.

8. Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.

    ![Didacticiel Excel-Figer l’en-tête](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a>Protéger une feuille de calcul

Dans cette étape du didacticiel, vous allez ajouter un autre bouton au ruban qui, lorsque l’utilisateur clique dessus, exécute une fonction qui vous allez définir et qui active/désactive la protection de la feuille de calcul.

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>Configuration du manifeste pour ajouter un deuxième bouton de ruban

1. Ouvrez le fichier manifeste my-office-add-in-manifest.xml.

2. Recherchez l’élément `<Control>`. Cet élément définit le bouton **Afficher le volet des pages** sur le ruban **Accueil** que vous utilisez pour lancer le complément. Nous allons ajouter un deuxième bouton au même groupe sur le ruban **Accueil**. Entre la balise Control de fin (`</Control>`) et la balise Group de fin (`</Group>`), ajoutez le balisage suivant.

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. Remplacez `TODO1` par une chaîne qui attribue un ID unique au bouton au sein de ce fichier manifeste. Étant donné que notre bouton va activer ou désactiver la protection de la feuille de calcul, utilisez « ToggleProtection ». Lorsque vous avez terminé, la balise Control de début doit ressembler à ce qui suit :

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. Les trois éléments `TODO` suivants définissent les éléments « resid », c’est-à-dire les ID de ressource. Une ressource est une chaîne. Vous allez créer ces trois chaînes lors d’une étape ultérieure. Pour l’instant, vous devez attribuer des ID aux ressources. L’étiquette du bouton doit indiquer « Toggle Protection », mais l’*ID* de cette chaîne doit être « ProtectionButtonLabel », donc l’élément `Label` terminé doit ressembler au code suivant :

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. L’élément `SuperTip` définit l’info-bulle du bouton. Le titre de l’info-bulle doit être identique à l’étiquette du bouton, nous utilisons donc le même ID de ressource : « ProtectionButtonLabel ». La description de l’info-bulle sera « Cliquez pour activer/désactiver la protection de la feuille de calcul ». Néanmoins, l’élément `ID` doit être « ProtectionButtonToolTip ». Ainsi, lorsque vous avez terminé, l’ensemble du balisage `SuperTip` doit ressembler au code suivant : 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > Dans un complément de production, vous n’utiliseriez pas la même icône pour deux boutons différents, mais pour simplifier ce didacticiel, nous allons le faire. Par conséquent, le balisage `Icon` de notre nouvel élément `Control` est simplement une copie de l’élément `Icon` provenant de l’élément `Control` existant. 

6. Le type de l’élément `Action` se trouvant à l’intérieur de l’élément `Control` d’origine qui était déjà présent dans le fichier manifeste est défini sur `ShowTaskpane`, mais notre nouveau bouton ne va pas ouvrir un volet Office, il va exécuter une fonction personnalisée que vous allez créer à une étape ultérieure. Il faut donc remplacer `TODO5` par `ExecuteFunction`, c’est-à-dire le type d’action pour les boutons qui déclenchent des fonctions personnalisées. La balise `Action` de début doit ressembler au code suivant :
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. L’élément `Action` d’origine possède des éléments enfants qui spécifient un ID de volet Office ainsi qu’une URL de la page qui doit être ouverte dans le volet Office. Toutefois, un élément `Action` de type `ExecuteFunction` comporte un élément enfant unique qui nomme la fonction que le contrôle exécute. Vous créerez cette fonction à une étape ultérieure, et la nommerez `toggleProtection`. Par conséquent, remplacez `TODO6` par le balisage suivant :
 
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
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
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

1. Ouvrez le fichier \function-file\function-file.js.

2. Le fichier possède déjà une expression de fonction appelée immédiatement (IIFE). *En dehors de la IIFE*, ajoutez le code suivant. Notez que nous spécifions un paramètre `args` pour la méthode et que la toute dernière ligne de la méthode appelle `args.completed`. Il s’agit d’une condition requise pour toutes les commandes de type **ExecuteFunction**. Elle signale à l’application hôte Office que la fonction est terminée et que l’interface utilisateur est à nouveau réactive.

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

3. Remplacez `TODO1` par le code suivant. Ce code utilise la propriété de protection de l’objet de feuille de calcul dans un modèle de bouton bascule standard. L’élément `TODO2` sera expliqué dans la section suivante.

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Ajoutez du code pour récupérer des propriétés de document dans les objets de script du volet Office

Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour écrire (*write*) dans le document Office. Chaque fonction se terminait par un appel de la méthode `context.sync()` qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées. Cependant, le code que vous avez ajouté dans la dernière étape appelle la propriété `sheet.protection.protected` et c’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `sheet` est uniquement un objet de proxy qui existe dans le script de votre volet Office. Il ne connaît pas l’état de protection réel du document, donc sa propriété `protection.protected` ne peut pas contenir une valeur réelle. Tout d’abord, il faut récupérer l’état de protection dans le document et l’utiliser pour définir la valeur de `sheet.protection.protected`. Seulement ensuite, la propriété `sheet.protection.protected` peut être appelée sans générer d’exception. Ce processus de récupération comporte trois étapes :

   1. Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.

   2. Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.

   3. Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.

Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.

1. Dans la fonction `toggleProtection`, remplacez `TODO2` par le code suivant. Remarque :
   
   - Chaque objet Excel possède une méthode `load`. Vous spécifiez les propriétés de l’objet que vous voulez lire dans le paramètre en tant que chaîne de noms séparés par des virgules. Dans ce cas, la propriété que vous devez lire est une sous-propriété de la propriété `protection`. Pour référence la sous-propriété, procédez presque exactement de la même façon que vous le feriez à n’importe quel autre emplacement de votre code, sauf que vous devez utiliser une barre oblique (« / ») au lieu d’un point « . ».

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

### <a name="configure-the-script-loading-html-file"></a>Configuration du fichier HTML de chargement de script

Ouvrez le fichier /function-file/function-file.html. Il s’agit d’un fichier HTML sans interface utilisateur qui est appelé lorsque l’utilisateur appuie sur le bouton **Toggle Worksheet Protection**. Son objectif consiste à charger la méthode JavaScript qui doit s’exécuter lorsque l’utilisateur appuie sur le bouton. Vous n’allez pas modifier ce fichier. Remarquez simplement que la deuxième balise `<script>` charge le fichier functionfile.js.

   > [!NOTE]
   > Le fichier function-file.html et le fichier function-file.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement distinct dans le volet Office du complément. Si le fichier function-file.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre l’objectif de groupement. En outre, le fichier function-file.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code. C’est pour ces deux raisons que ce complément ne transpile pas le fichier function-file.js du tout. 

### <a name="test-the-add-in"></a>Test du complément

1. Fermez toutes les applications Office, y compris Excel. 

2. Supprimez le cache Office en supprimant le contenu du dossier de cache. Cette opération est nécessaire pour effacer complètement de l’hôte l’ancienne version du complément. 

    - Pour Windows : `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

    - Pour Mac : `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

3. Si, pour une quelconque raison, votre serveur n’est pas en cours d’exécution, accédez au dossier **Start** du projet et exécutez la commande `npm start` dans une fenêtre Git Bash ou une invite système Node.JS. Vous n’avez pas besoin de recréer le projet, car le seul fichier JavaScript que vous avez modifié ne fait pas partie du fichier bundle.js créé.

4. À l’aide de la nouvelle version du fichier manifeste modifié, répétez le processus de chargement de version test en utilisant l’une des méthodes suivantes. *Vous devez remplacer la copie précédente du fichier manifeste.*

    - Windows : [Chargement de version test des compléments Office](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)

    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

5. Ouvrez une feuille de calcul dans Excel.

6. Sur le ruban **Accueil**, sélectionnez **Toggle Worksheet Protection** (Activer/Désactiver la protection de la feuille de calcul). Notez que la plupart des contrôles figurant sur le ruban sont désactivés (et visuellement grisés) comme illustré dans la capture d’écran ci-dessous. 

7. Sélectionnez une cellule comme vous le feriez si vous vouliez modifier son contenu. Vous rencontrez une erreur indiquant que la feuille de calcul est protégée.

8. Sélectionnez **Toggle Worksheet Protection** à nouveau pour réactiver les contrôles. Vous pouvez alors modifier une nouvelle fois les valeurs de cellule.

    ![Didacticiel Excel-Ruban avec protection activée](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a>Ouvrir une boîte de dialogue

Dans cette étape finale du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de boîte de dialogue au processus de volet Office et fermer la boîte de dialogue. Les boîtes de dialogue des compléments Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page hôte dans le volet Office.

### <a name="create-the-dialog-page"></a>Création de la page de boîte de dialogue

1. Ouvrez le projet dans votre éditeur de code.

2. Créez un fichier à la racine du projet (où se trouve le fichier index.html) et nommez-le popup.html.

3. Ajoutez le balisage suivant au fichier popup.html. Remarque :

   - La page comporte un champ `<input>`, dans lequel l’utilisateur entrera son nom, et un bouton qui permet d’envoyer le nom à la page dans le volet Office où il sera affiché.

   - Le balisage charge un script appelé popup.js que vous allez créer dans une étape ultérieure.

   - Il charge également la bibliothèque Office.JS et jQuery, car ils seront utilisés dans popup.js.

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. Créez un fichier à la racine du projet et nommez-le popup.js.

5. Ajoutez le code suivant au fichier popup.js. Tenez compte des informations suivantes :

   - *Toutes les pages appellent les API dans la bibliothèque Office.JS doivent tout d’abord vérifier que la bibliothèque est entièrement initialisée.* La meilleure façon de procéder consiste à appeler la méthode `Office.onReady()`. Si votre complément dispose de ses propres tâches d’initialisation, le code doit passer dans une méthode `then()` chaînée à l’appel de `Office.onReady()`. Pour voir un exemple, affichez le fichier app.js à la racine du projet. Le code qui appelle `Office.onReady()` doit être exécuté avant tout appel à Office.JS ; l’affectation se trouve donc dans un fichier de script chargé par la page, comme dans ce cas.
   - La fonction `ready` jQuery est appelée à l’intérieur de la méthode `then()`. Dans la plupart des cas, le code de chargement, d’initialisation ou d’amorçage d’autres bibliothèques JavaScript se trouve à l’intérieur de la méthode `then()` chaînée à l’appel de `Office.onReady()`.

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {
                $(document).ready(function () {  

                    // TODO1: Assign handler to the OK button.

                });
            });

        // TODO2: Create the OK button handler

    }());
    ```

6. Remplacez `TODO1` par le code suivant. Vous allez créer la fonction `sendStringToParentPage` à l’étape suivante.

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. Remplacez `TODO2` par le code suivant. La méthode `messageParent` transmet son paramètre à la page parent, qui est, dans ce cas, la page dans le volet Office. Le paramètre peut être une valeur booléenne ou une chaîne qui inclut tous les éléments qui peuvent être sérialisés en tant que chaîne, au format XML ou JSON.

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. Enregistrez le fichier.

   > [!NOTE]
   > Le fichier popup.html et le fichier popup.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement séparé à partir du volet Office du complément. Si le popup.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre de l’objectif de groupement. En outre, le fichier popup.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code. C’est pour ces deux raisons que ce complément ne transpile pas le fichier popup.js du tout.

### <a name="open-the-dialog-from-the-task-pane"></a>Ouverture de la boîte de dialogue à partir du volet Office

1. Ouvrez le fichier index.html.

2. Sous la balise `div` qui contient le bouton `freeze-header`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. La boîte de dialogue invitera l’utilisateur à saisir son nom et transmettra ce nom au volet Office. Le volet Office s’affichera dans une étiquette. Juste en dessous de la balise `div` que vous venez d’ajouter, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Sous la ligne qui attribue un gestionnaire de clics au bouton `freeze-header`, ajoutez le code suivant. Vous allez créer la méthode `openDialog` à une étape ultérieure.

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. Ajoutez la déclaration suivante sous la fonction `freezeHeader`. Cette variable est utilisée pour conserver un objet dans le contexte d’exécution de la page parent qui agit en tant qu’intermédiaire pour le contexte d’exécution de la page de boîte de dialogue.

    ```js
    var dialog = null;
    ```

7. Sous la déclaration de la balise `dialog`, ajoutez la fonction suivante. Le plus important à remarquer à propos de ce code est ce qui ne s’y trouve *pas* : il n’y a aucun appel de `Excel.run`. Cela est dû au fait que l’API d’ouverture de boîte de dialogue est partagée par tous les hôtes Office, elle fait donc partie de l’API commune JavaScript Office, pas de l’API spécifique d’Excel.

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

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

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Traitement du message à partir de la boîte de dialogue et fermeture de la boîte de dialogue

1. Continuez dans le fichier app.js et remplacez `TODO2` par le code suivant. Remarque :

   - Le rappel est exécuté immédiatement après que la boîte de dialogue s’est ouverte correctement et avant que l’utilisateur ait pris une quelconque action dans la boîte de dialogue.

   - `result.value` représente l’objet qui agit comme un intermédiaire entre les contextes d’exécution des pages parent et de boîte de dialogue.

   - La fonction `processMessage` sera créée à une étape ultérieure. Ce gestionnaire traitera toutes les valeurs envoyées par la page de boîte de dialogue avec les appels de la fonction `messageParent`.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. Sous la fonction `openDialog`, ajoutez la fonction suivante.

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur **Ctrl+C **à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.

5. Sélectionnez le bouton **Boîte de dialogue Ouvrir** dans le volet Office.

6. Lorsque la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la. Notez que vous pouvez interagir avec la feuille de calcul, appuyez sur les autres boutons dans le volet Office, mais vous ne pouvez pas lancer une deuxième boîte de dialogue à partir de la même page de volet de tâches.

7. Dans la boîte de dialogue, entrez un nom et appuyez sur **OK**. Ce nom apparaît sur le volet Office et la boîte de dialogue se ferme.

8. Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction `processMessage`. Ensuite, répétez les étapes de cette section. La boîte de dialogue reste ouverte et vous pouvez modifier le nom. Vous pouvez la fermer manuellement en appuyant sur la croix (**X**) en haut à droite.

    ![Didacticiel Excel- Boîte de dialogue](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a>Étapes suivantes

Ce didacticiel vous apprend à créer un complément Excel qui interagit avec des tableaux, des graphiques (chart), des feuilles de calcul et des boîtes de dialogue dans un classeur Excel. Pour en savoir plus sur le développement des complément Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Présentation des compléments Excel](../excel/excel-add-ins-overview.md)
