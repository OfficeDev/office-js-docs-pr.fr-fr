Dans cette étape du didacticiel, vous devez tester par programme que votre complément prend en charge la version actuelle de Word de l’utilisateur, puis insérer un paragraphe dans le document.

> [!NOTE]
> Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="code-the-add-in"></a>Codage du complément

1. Ouvrez le projet dans votre éditeur de code. 
2. Ouvrez le fichier index.html.
3. Remplacez `TODO1` par le codage suivant :

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. Ouvrez le fichier app.js.
5. Remplacez `TODO1` par le code suivant. Ce code détermine si la version de Word de l’utilisateur prend en charge une version de Word.js qui inclut toutes les API utilisées dans les étapes de ce didacticiel. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge. Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par sa version d’Excel.

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. Remplacez `TODO2` par le code suivant :

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Remplacez `TODO3` par le code suivant : Remarques :
   - Votre logique métier Word.js est ajoutée à la fonction qui est transmise à `Word.run`. Cette logique n’est pas exécutée immédiatement. Au lieu de cela, elle est ajoutée à une file d’attente de commandes.
   - La méthode `context.sync` envoie toutes les commandes en file d’attente vers Word pour exécution.
   - L’élément `Word.run` est suivi par un bloc `catch`. Il s’agit d’une meilleure pratique que vous devez toujours suivre. 

    ```js
    function insertParagraph() {
        Word.run(function (context) {
            
            // TODO4: Queue commands to insert a paragraph into the document.

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

8. Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :
   - Le premier paramètre de la méthode `insertParagraph` correspond au texte pour le nouveau paragraphe.
   - Le deuxième paramètre correspond à l’emplacement dans le corps où sera inséré le paragraphe. Les autres options d’insertion de paragraphe, lorsque l’objet parent est le corps, sont « Fin » et « Remplacer ». 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a>Test du complément

1. Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.
2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.
3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.   
4. Chargez une version test du complément en utilisant l’une des méthodes suivantes :
    - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online : [Chargement d’une version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. Dans le menu **Accueil** de Word, sélectionnez **Afficher le volet des tâches**.
6. Dans le volet des tâches, sélectionnez **Insérer un paragraphe**.
7. Apportez une modification au paragraphe. 
8. Sélectionnez à nouveau **Insérer un paragraphe**. Notez que le nouveau paragraphe se trouve au-dessus du paragraphe précédent, car la méthode `insertParagraph` effectue l’insertion au « début » du corps du document.

    ![Didacticiel Word - Insérer un paragraphe](../images/word-tutorial-insert-paragraph.png)
