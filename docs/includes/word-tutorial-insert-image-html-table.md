Dans cette étape du didacticiel, vous allez découvrir comment insérer des images, du code HTML et des tableaux dans le document.

> [!NOTE]
> Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="insert-an-image"></a>Insérer une image

1. Ouvrez le projet dans votre éditeur de code. 
2. Ouvrez le fichier index.html.
3. En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-image">Insert Image</button>            
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Dans la partie supérieure du fichier, juste en dessous de la ligne stricte, ajoutez la ligne suivante. Cette ligne importe une variable à partir d’un autre fichier. La variable est une chaîne en base 64 qui encode une image. Pour afficher la chaîne encodée, ouvrez le fichier base64Image.js dans la racine du projet.

    ```js
    import { base64Image } from "./base64Image";
    ``` 

5. Sous la ligne qui attribue un gestionnaire de clics au bouton `replace-text`, ajoutez le code suivant :

    ```js
    $('#insert-image').click(insertImage);
    ```

6. Sous la fonction `replaceText`, ajoutez la fonction suivante :

    ```js
    function insertImage() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert an image.

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

7. Remplacez `TODO1` par le code suivant. Cette ligne insère l’image encodée en base 64 à la fin du document. (L’objet `Paragraph` contient également une méthode `insertInlinePictureFromBase64` et d’autres méthodes `insert*`. Reportez-vous à la section Insérer du code HTML suivante pour consulter un exemple.)

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ``` 

## <a name="insert-html"></a>Insérer du code HTML

1. Ouvrez le fichier index.html.
2. En dessous de la balise `div` qui contient le bouton `insert-image`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-html">Insert HTML</button>            
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-image`, ajoutez le code suivant :

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. Sous la fonction `insertImage`, ajoutez la fonction suivante :

    ```js
    function insertHTML() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a string of HTML.

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

6. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :
   - La première ligne ajoute un paragraphe vide à la fin du document. 
   - La deuxième ligne insère une chaîne de code HTML à la fin du paragraphe. Plus précisément, deux paragraphes : un paragraphe avec la police Verdana, et l’autre avec le style par défaut du document Word. (Comme pour la méthode `insertImage` précédente, l’objet `context.document.body` contient également les méthodes `insert*`.)

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ``` 

## <a name="insert-table"></a>Insérer un tableau

1. Ouvrez le fichier index.html.
3. En dessous de la balise `div` qui contient le bouton `insert-html`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-table">Insert Table</button>            
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-html`, ajoutez le code suivant :

    ```js
    $('#insert-table').click(insertTable);
    ```

6. Sous la fonction `insertHTML`, ajoutez la fonction suivante :

    ```js
    function insertTable() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

7. Remplacez `TODO1` par le code suivant. Cette ligne utilise la méthode `ParapgraphCollection.getFirst` pour obtenir une référence au premier paragraphe, puis utilise la méthode `Paragraph.getNext` pour obtenir une référence au deuxième paragraphe.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ``` 

8. Remplacez `TODO2` par le code suivant. Veuillez noter les informations suivantes :
   - Les deux premiers paramètres de la méthode `insertTable` spécifient le nombre de lignes et de colonnes.
   - Le troisième paramètre indique l’emplacement où insérer le tableau, en l’occurrence après le paragraphe.
   - Le quatrième paramètre est une matrice à deux dimensions qui définit les valeurs des cellules du tableau.
   - Le tableau aura un style par défaut brut, mais la méthode `insertTable` renvoie un objet `Table` avec de nombreux membres, dont certains sont utilisés pour définir le style du tableau.

     ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ``` 

## <a name="test-the-add-in"></a>Test du complément


1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.
3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.
4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.
5. Dans le volet des tâches, sélectionnez **Insérer un paragraphe** au moins trois fois pour vous assurer qu’il existe quelques paragraphes dans le document.
6. Sélectionnez le bouton **Insérer une image** et notez qu’une image est insérée à la fin du document.
7. Sélectionnez le bouton **Insérer du code HTML**, puis notez que deux paragraphes sont insérés à la fin du document, et que le premier est affiché dans la police Verdana.
8. Sélectionnez le bouton **Insérer un tableau** et notez qu’un tableau est inséré après le deuxième paragraphe.

    ![Didacticiel Word - Insérer une image, du code HTML et un tableau](../images/word-tutorial-insert-image-html-table.png)
