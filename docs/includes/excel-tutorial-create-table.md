Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.

> [!NOTE]
> Cette page décrit une étape individuelle du didacticiel sur le complément Excel. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="code-the-add-in"></a>Codage du complément

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

7. Remplacez `TODO3` par le code suivant : Remarques :
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

8. Remplacez `TODO4` par le code suivant. Remarque :
   - Le code crée un tableau à l’aide de la méthode `add` de collection de tableau d’une feuille de calcul, qui existe toujours même si elle est vide. Il s’agit de la méthode standard de création d’objets Excel.js. Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel. Au lieu de cela, vous l’ajoutez à un objet de la collection parent.
   - Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau. La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes. Il s’agit d’un modèle plus courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.
   - Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
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

## <a name="test-the-add-in"></a>Test du complément

1. Ouvrez une fenêtre Git Bash ou une invite système activée par Node.JS, et accédez au dossier **Démarrer** du projet.
2. Exécutez la commande `npm run build` pour transpiler votre code source ES6 sur une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisée en arrière-plan par Excel pour exécuter les compléments Excel).
3. Exécutez la commande `npm start` pour démarrer un serveur web exécuté sur un hôte local.
4. Chargez une version test du complément en utilisant l’une des méthodes suivantes :
    - Windows : [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad et Mac : [Chargement de versions test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. Dans le menu **Accueil**, sélectionnez **Afficher le volet Office**.
6. Dans le volet Office, sélectionnez **Créer un tableau**.

    ![Didacticiel Excel - Créer un tableau](../images/excel-tutorial-create-table.png)
