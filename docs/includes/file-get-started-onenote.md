# <a name="build-your-first-onenote-add-in"></a>Cr?er votre premier compl?ment OneNote

Cet article d?crit le processus de cr?ation d?un compl?ment OneNote ? l?aide de jQuery et de l?API JavaScript pour Office.

## <a name="prerequisites"></a>Conditions pr?alables

- [Node.js](https://nodejs.org)

- Installez la derni?re version de [Yeoman](https://github.com/yeoman/yo) et le [g?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>Cr?ation du projet de compl?ment

1. Cr?ez un dossier sur votre lecteur local et nommez-le `my-onenote-addin`. Il s?agit de l?emplacement dans lequel vous allez cr?er les fichiers de votre compl?ment.

2. Acc?dez ? votre nouveau dossier.

    ```bash
    cd my-onenote-addin
    ```

3. Utilisez le g?n?rateur Yeoman afin de cr?er un projet de compl?ment OneNote. Ex?cutez la commande suivante, puis r?pondez aux invites comme suit :

    ```bash
    yo office
    ```

    - **Voulez-vous cr?er un sous-dossier de votre projet ? :** `No`
    - **Comment souhaitez-vous nommer votre compl?ment ? :** `OneNote Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `OneNote`
    - **Voulez-vous cr?er un compl?ment ? :** `Yes`
    - **Souhaitez-vous utiliser TypeScript ? :** `No`
    - **Choisissez une infrastructure :** `Jquery`

    Le g?n?rateur demande ensuite si vous voulez ouvrir **resource.html**. Il n?est pas n?cessaire de l?ouvrir pour ce didacticiel, mais n?h?sitez pas ? l?ouvrir si vous ?tes curieux. Cliquez sur Oui ou Non pour fermer l?assistant et laisser le g?n?rateur faire son travail.

    ![Capture d??cran des invites et des r?ponses relatives au g?n?rateur Yeoman](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a>Mise ? jour du code

1. Dans votre ?diteur de code, ouvrez **index.html** ? la racine du projet. Ce fichier contient le code HTML qui s?affichera dans le volet Office du compl?ment.

2. Remplacez l??l?ment `<main>` dans l??l?ment `<body>` par le balisage suivant et enregistrez le fichier. Cette option ajoute une zone de texte et un bouton ? l?aide des [composants de la structure de l?interface utilisateur d?Office](http://dev.office.com/fabric/components).

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

3. Ouvrez le fichier **app.js** pour sp?cifier le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

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

## <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier nomm? **one-note-add-in-manifest.xml** pour d?finir les param?tres et les fonctionnalit?s du compl?ment.

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for OneNote**.

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

## <a name="start-the-dev-server"></a>D?marrage du serveur de d?veloppement

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>Essayez !

1. Dans [OneNote Online](https://www.onenote.com/notebooks), ouvrez un bloc-notes.

2. Choisissez **Insertion > Compl?ments Office** pour ouvrir la bo?te de dialogue Compl?ments Office.

    - Si vous ?tes connect? avec votre compte de consommateur, s?lectionnez l?onglet **MES COMPL?MENTS**, puis choisissez **T?l?charger mon compl?ment**.

    - Si vous ?tes connect? avec votre compte professionnel ou scolaire, s?lectionnez l?onglet **MON ORGANISATION**, puis choisissez **T?l?charger mon compl?ment**. 

    L?image suivante montre l?onglet **MES COMPL?MENTS** pour les blocs-notes de consommateurs.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. Dans la bo?te de dialogue T?l?charger le compl?ment, acc?dez ? **one-note-add-in-manifest.xml** dans le dossier de projet, puis choisissez **T?l?charger**. 

4. Depuis l?onglet **Accueil**, cliquez le bouton **Afficher le volet Office** du ruban. Le compl?ment volet Office s?ouvre dans un iFrame ? c?t? de la page OneNote.

5. Entrez du texte dans la zone de texte, puis choisissez **Ajouter un plan**. Le texte que vous avez entr? est ajout? ? la page. 

    ![Compl?ment OneNote g?n?r? ? partir de cette proc?dure pas ? pas](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a>Conseils et r?solution des probl?mes

- Vous pouvez d?boguer le compl?ment ? l?aide des outils de d?veloppement de votre navigateur. Lorsque vous utilisez le serveur web Gulp et le d?bogage dans Internet Explorer ou Chrome, vous pouvez enregistrer les modifications localement et simplement actualiser l?iFrame du compl?ment.

- Lorsque vous examinez un objet OneNote, les propri?t?s qui sont actuellement disponibles affichent les valeurs r?elles. Les propri?t?s qui doivent ?tre charg?es sont affich?es comme *non d?finies*. D?veloppez le n?ud `_proto_` pour visualiser les propri?t?s qui sont d?finies sur l?objet, mais qui ne sont pas encore charg?es.

   ![Objet OneNote d?charg? dans le d?bogueur](../images/onenote-debug.png)

- Vous devez activer le contenu mixte dans le navigateur si votre compl?ment utilise des ressources HTTP. Les compl?ments de production doivent uniquement utiliser des ressources HTTPS s?curis?es.

- Les compl?ments de volet Office peuvent ?tre ouverts ? partir de n?importe o?, mais les compl?ments de contenu peuvent uniquement ?tre ins?r?s ? l?int?rieur de contenu de page normal (et non dans des titres, des images, des iFrames, etc.). 

## <a name="next-steps"></a>?tapes suivantes

F?licitations, vous avez cr?? un compl?ment OneNote ! Ensuite, vous allez ?tudier en d?tail les concepts fondamentaux de la cr?ation de compl?ments Excel.

> [!div class="nextstepaction"]
> [Vue d?ensemble de la programmation de l?API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>Voir aussi

- [Vue d?ensemble de la programmation de l?API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [R?f?rence de l?API JavaScript de OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Exemple de grille d??valuation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d?ensemble de la plateforme des compl?ments Office](../overview/office-add-ins.md)
