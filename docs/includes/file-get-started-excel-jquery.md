# <a name="build-an-excel-add-in-using-jquery"></a>D?veloppement d?un compl?ment Excel ? l?aide de jQuery

Cet article d?crit le processus de cr?ation d?un compl?ment Excel ? l?aide de jQuery et de l?API JavaScript pour Excel. 

## <a name="create-the-add-in"></a>Cr?er le compl?ment 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[Visual Studio](#tab/visual-studio)

### <a name="prerequisites"></a>Conditions pr?alables

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Cr?ation du projet de compl?ment

1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.
    
2. Dans la liste des types de projet, sous **Visual C#** ou **Visual Basic**, d?veloppez **Office/SharePoint**, choisissez **Compl?ments**, puis **Compl?ment Excel Web** pour le type de projet. 

3. Nommez le projet, puis cliquez sur **OK**.

4. Dans la fen?tre de dialogue **Cr?er un compl?ment Office**, s?lectionnez **Ajouter de nouvelles fonctionnalit?s ? Excel**, puis s?lectionnez **Terminer** pour cr?er le projet.

5. Visual Studio cr?e une solution et ses deux projets apparaissent dans l?**explorateur de solutions**. Le fichier **Home.html** s?ouvre dans Visual Studio.
    
### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>Mise ? jour du code

1. **Home.html** sp?cifie le code HTML qui s?affichera dans le volet Office du compl?ment. Dans **Home.html**, remplacez l??l?ment `<body>` par le balisage suivant et enregistrez le fichier.
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. Ouvrez le fichier **Home.js** ? la racine du projet d?application web. Ce fichier sp?cifie le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier. 

    ```js
    'use strict';

    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. Ouvrez le fichier **Home.css** ? la racine du projet d?application web. Ce fichier sp?cifie les styles personnalis?s pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier. 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier manifeste XML dans le projet de compl?ment. Ce fichier d?finit les param?tres et les fonctionnalit?s du compl?ment.

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `DisplayName` poss?de un espace r?serv?. Remplacez-le par **My Office Add-in**.

4. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for Excel**.

5. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a>Essayez !

1. ? l?aide de Visual Studio, testez le nouveau compl?ment Excel en appuyant sur F5 ou en choisissant le bouton **D?marrer** pour lancer Excel avec le bouton du compl?ment **Show Taskpane** (Afficher le volet Office) qui appara?t dans le ruban. Le compl?ment sera h?berg? localement sur IIS.

2. Dans Excel, s?lectionnez l?onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du compl?ment.

    ![Bouton Compl?ment Excel](../images/excel-quickstart-addin-2a.png)

3. S?lectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **D?finir couleur** pour d?finir la couleur de la plage s?lectionn?e en vert.

    ![Compl?ment Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[Tous les ?diteurs](#tab/visual-studio-code)

### <a name="prerequisites"></a>Conditions pr?alables

- [Node.js](https://nodejs.org)

- Installez la derni?re version de [Yeoman](https://github.com/yeoman/yo) et le [g?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a>Cr?ation de l?application web

1. Cr?ez un dossier sur votre lecteur local et nommez-le **my-addin**. Il s?agit de l?endroit o? vous allez cr?er les fichiers de votre application.

2. Acc?dez au dossier de votre application.

    ```bash
    cd my-addin
    ```

3. Utilisez le g?n?rateur Yeoman pour g?n?rer le fichier manifeste de votre compl?ment. Ex?cutez la commande suivante, puis r?pondez aux invites comme indiqu? dans la capture d??cran suivante :

    ```bash
    yo office
    ```

    - **Voulez-vous cr?er un sous-dossier de votre projet ? :** `No`
    - **Comment souhaitez-vous nommer votre compl?ment ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `Excel`
    - **Voulez-vous cr?er un compl?ment ? :** `Yes`
    - **Souhaitez-vous utiliser TypeScript ? :** `No`
    - **Choisissez une infrastructure :** `Jquery`

    Le g?n?rateur demande ensuite si vous voulez ouvrir **resource.html**. Il n?est pas n?cessaire de l?ouvrir pour ce didacticiel, mais n?h?sitez pas ? l?ouvrir si vous ?tes curieux. Cliquez sur Oui ou Non pour fermer l?assistant et laisser le g?n?rateur faire son travail.

    ![G?n?rateur Yeoman](../images/yo-office-jquery.png)


4. Dans votre ?diteur de code, ouvrez **index.html** ? la racine du projet. Ce fichier sp?cifie le code HTML qui s?affichera dans le volet Office du compl?ment. 
 
5. Dans **index.html**, remplacez la balise `header` g?n?r?e par le balisage suivant.
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. Dans **index.html**, remplacez la balise `main` g?n?r?e par le balisage suivant et enregistrez le fichier.

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. Ouvrez le fichier **app.js** pour sp?cifier le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    'use strict';
    
    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

8. Ouvrez le fichier **app.css** pour sp?cifier les styles personnalis?s pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier nomm? **my-office-add-in-manifest.xml** pour d?finir les param?tres et les fonctionnalit?s du compl?ment. 

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `DisplayName` poss?de un espace r?serv?. Remplacez-le par **My Office Add-in**.

4. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for Excel**.

5. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a>D?marrage du serveur de d?veloppement

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez afin d?ex?cuter votre compl?ment en vue d?en charger une version test dans Excel.

    - Windows : [Chargement de version test des compl?ments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de versions test des compl?ments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compl?ments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Dans Excel, s?lectionnez l?onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du compl?ment.

    ![Bouton Compl?ment Excel](../images/excel-quickstart-addin-2b.png)

3. S?lectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **D?finir couleur** pour d?finir la couleur de la plage s?lectionn?e en vert.

    ![Compl?ment Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a>?tapes suivantes

F?licitations, vous avez cr?? un compl?ment Excel ? l?aide de jQuery ! D?couvrez ? pr?sent les fonctionnalit?s des compl?ments Excel et cr?ez un compl?ment plus complexe en continuant le didacticiel sur le compl?ment Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compl?ments Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compl?ments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts de base de l?API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compl?ments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [R?f?rence de l?API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
