# <a name="build-your-first-onenote-add-in"></a>Créer votre premier complément OneNote

Cet article décrit le processus de création d’un complément OneNote à l’aide de jQuery et de l’API JavaScript pour Office.

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org)

- Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>Création du projet de complément

1. Créez un dossier sur votre lecteur local et nommez-le `my-onenote-addin`. Il s’agit de l’emplacement dans lequel vous allez créer les fichiers de votre complément.

2. Accédez à votre nouveau dossier.

    ```bash
    cd my-onenote-addin
    ```

3. Utilisez le générateur Yeoman afin de créer un projet de complément OneNote. Exécutez la commande suivante, puis répondez aux invites comme suit :

    ```bash
    yo office
    ```

    - **Choisissez un type de projet :** `Jquery`
    - **Choisissez un type de script :** `Javascript`
    - **Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `Onenote`

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-onenote-jquery.png)
    
    Une fois que vois avez terminé avec l'assistant, le générateur crée le projet et installe les composants Node de prise en charge.


## <a name="update-the-code"></a>Mise à jour du code

1. Dans votre éditeur de code, ouvrez **index.html** à la racine du projet. Ce fichier contient le code HTML qui s’affichera dans le volet Office du complément.

2. Remplacez l’élément `<main>` dans l’élément `<body>` par le balisage suivant et enregistrez le fichier. Cette option ajoute une zone de texte et un bouton à l’aide des [composants de la structure de l’interface utilisateur d’Office](http://dev.office.com/fabric/components).

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. Ouvrez le fichier **src\index.js** pour spécifier le script pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a>Mise à jour du manifeste

1. Ouvrez le fichier nommé **one-note-add-in-manifest.xml** pour définir les paramètres et les fonctionnalités du complément.

2. L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.

3. L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for OneNote**.

4. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Démarrage du serveur de développement

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>Essayez !

1. Dans [OneNote Online](https://www.onenote.com/notebooks), ouvrez un bloc-notes.

2. Choisissez **Insertion > Compléments Office** pour ouvrir la boîte de dialogue Compléments Office.

    - Si vous êtes connecté avec votre compte de consommateur, sélectionnez l’onglet **MES COMPLÉMENTS**, puis choisissez **Télécharger mon complément**.

    - Si vous êtes connecté avec votre compte professionnel ou scolaire, sélectionnez l’onglet **MON ORGANISATION**, puis choisissez **Télécharger mon complément**. 

    L’image suivante montre l’onglet **MES COMPLÉMENTS** pour les blocs-notes de consommateurs.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. Dans la boîte de dialogue Télécharger le complément, accédez à **one-note-add-in-manifest.xml** dans le dossier de projet, puis choisissez **Télécharger**. 

4. Dans l’onglet **Accueil,** cliquez sur le bouton **Afficher le volet des tâches** du ruban. Le complément volet Office s’ouvre dans un iFrame à côté de la page OneNote.

5. Entrez du texte dans la zone de texte, puis choisissez **Ajouter un contour**. Le texte que vous avez entré est ajouté à la page. 

    ![Complément OneNote généré à partir de cette procédure pas à pas](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a>Conseils et résolution des problèmes

- Vous pouvez déboguer le complément à l’aide des outils de développement de votre navigateur. Lorsque vous utilisez le serveur web Gulp et le débogage dans Internet Explorer ou Chrome, vous pouvez enregistrer les modifications localement et simplement actualiser l’iFrame du complément.

- Lorsque vous examinez un objet OneNote, les propriétés qui sont actuellement disponibles affichent les valeurs réelles. Les propriétés qui doivent être chargées sont affichées comme *non définies*. Développez le nœud `_proto_` pour visualiser les propriétés qui sont définies sur l’objet, mais qui ne sont pas encore chargées.

   ![Objet OneNote déchargé dans le débogueur](../images/onenote-debug.png)

- Vous devez activer le contenu mixte dans le navigateur si votre complément utilise des ressources HTTP. Les compléments de production doivent uniquement utiliser des ressources HTTPS sécurisées.

- Les compléments de volet Office peuvent être ouverts à partir de n’importe où, mais les compléments de contenu peuvent uniquement être insérés à l’intérieur de contenu de page normal (et non dans des titres, des images, des iFrames, etc.). 

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément OneNote ! Ensuite, vous allez étudier en détail les concepts fondamentaux de la création de compléments Excel.

> [!div class="nextstepaction"]
> [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Référence de l’API JavaScript de OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
