Dans cette étape finale du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de boîte de dialogue au processus de volet Office et fermer la boîte de dialogue. Les boîtes de dialogue des compléments Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page hôte dans le volet Office.

> [!NOTE]
> Cette page décrit une étape individuelle du didacticiel sur le complément Excel. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="create-the-dialog-page"></a>Création de la page de boîte de dialogue

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
5. Ajoutez le code suivant au fichier popup.js. Remarque :
   - *Toutes les pages qui appellent des API dans la bibliothèque Office.JS doivent affecter une fonction à la propriété `Office.initialize`.* Si aucune initialisation n’est nécessaire, la fonction peut avoir un corps vide, mais la propriété ne doit pas être laissée indéfinie, affectée à null ni à une valeur qui n’est pas une fonction. Pour voir un exemple, affichez le fichier app.js à la racine du projet. Le code qui exécute l’affectation doit être exécuté avant tout appel à Office.JS ; l’affectation se trouve donc dans un fichier de script chargé par la page, comme dans ce cas.
   - La fonction `ready` jQuery est appelée à l’intérieur de la méthode `initialize`. Une règle quasi-universelle veut que le code de chargement, d’initialisation ou d’amorçage des autres bibliothèques JavaScript se trouve à l’intérieur de la fonction `Office.initialize`.

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {
            $(document).ready(function () {  

                // TODO1: Assign handler to the OK button.

            });
        }

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


## <a name="open-the-dialog-from-the-task-pane"></a>Ouverture de la boîte de dialogue à partir du volet Office

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
    let dialog = null;
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

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Traitement du message à partir de la boîte de dialogue et fermeture de la boîte de dialogue

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

## <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Une fois la commande build exécutée, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

1. Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).
2. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.
4. Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.
6. Sélectionnez le bouton **Boîte de dialogue Ouvrir** dans le volet Office.
7. Lorsque la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la. Vous pouvez interagir avec la feuille de calcul et appuyer sur les autres boutons du volet Office. Pour autant, vous ne pouvez pas lancer une deuxième boîte de dialogue à partir de la même page de volet Office.
8. Dans la boîte de dialogue, entrez un nom et appuyez sur **OK**. Ce nom apparaît sur le volet Office et la boîte de dialogue se ferme.
9. Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction `processMessage`. Ensuite, répétez les étapes de cette section. La boîte de dialogue reste ouverte et vous pouvez modifier le nom. Vous pouvez la fermer manuellement en appuyant sur la croix (**X**) en haut à droite.

    ![Didacticiel Excel - Boîte de dialogue](../images/excel-tutorial-dialog-open.png)
