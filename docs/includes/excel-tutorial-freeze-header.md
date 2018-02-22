Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée. Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas. 

## <a name="freeze-the-tables-header-row"></a>Figer la ligne d’en-tête du tableau

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
   - La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être figées. Nous transmettons la valeur `1` pour épingler la première ligne.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 

## <a name="test-the-add-in"></a>Tester le complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris au fichier app.js, il ne retranspile pas le code JavaScript, donc vous devez répéter la commande build afin que vos modifications app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

1. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).
2. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.
4. Rechargez le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet Office** pour rouvrir le complément.
6. Si le tableau est dans la feuille de calcul, supprimez-le.
7. Dans le volet Office, sélectionnez **Créer un tableau**. 
8. Sélectionnez le bouton **Freeze Header**.
9. Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.

    ![Didacticiel Excel - Figer l’en-tête](../images/excel-tutorial-freeze-header.png)
