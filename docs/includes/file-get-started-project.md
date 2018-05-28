# <a name="build-your-first-project-add-in"></a>Cr?ation de votre premier compl?ment Project

Cet article d?crit le processus de cr?ation d?un compl?ment Project ? l?aide de jQuery et de l?API JavaScript pour Office.

## <a name="prerequisites"></a>Conditions pr?alables

- [Node.js](https://nodejs.org)

- Installez la derni?re version de [Yeoman](https://github.com/yeoman/yo) et le [g?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>Cr?er le compl?ment

1. Cr?ez un dossier sur votre lecteur local et nommez-le `my-project-addin`. Il s?agit de l?emplacement dans lequel vous allez cr?er les fichiers de votre compl?ment.

2. Acc?dez ? votre nouveau dossier.

    ```bash
    cd my-project-addin
    ```

3. Utilisez le g?n?rateur Yeoman afin de cr?er un projet de compl?ment Project. Ex?cutez la commande suivante, puis r?pondez aux invites comme suit :

    ```bash
    yo office
    ```

    - **Voulez-vous cr?er un sous-dossier de votre projet ? :** `No`
    - **Comment souhaitez-vous nommer votre compl?ment ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `Project`
    - **Voulez-vous cr?er un compl?ment ? :** `Yes`
    - **Souhaitez-vous utiliser TypeScript ? :** `No`
    - **Choisissez une infrastructure :** `Jquery`

    Le g?n?rateur demande ensuite si vous voulez ouvrir **resource.html**. Il n?est pas n?cessaire de l?ouvrir pour ce didacticiel, mais n?h?sitez pas ? l?ouvrir si vous ?tes curieux. Cliquez sur Oui ou Non pour fermer l?assistant et laisser le g?n?rateur faire son travail.

    ![Capture d??cran des invites et des r?ponses relatives au g?n?rateur Yeoman](../images/yo-office-project-jquery.png)

## <a name="update-the-code"></a>Mise ? jour du code

1. Dans votre ?diteur de code, ouvrez **index.html** ? la racine du projet. Ce fichier contient le code HTML qui s?affichera dans le volet Office du compl?ment.

2. Remplacez l??l?ment `<header>` ? l?int?rieur de l??l?ment `<body>` par le balisage suivant.

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. Remplacez l??l?ment `<main>` dans l??l?ment `<body>` par le balisage suivant et enregistrez le fichier.

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
            <h3>Try it out</h3>
            <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
            <br/><br/>
            <button class="ms-Button" id="get-task">Get Task data</button>
            <br/>
            <h4>Results:</h4>
            <textarea id="result" rows="6" cols="25"></textarea>
        </div>
    </div>
    ```

4. Ouvrez le fichier **app.js** pour sp?cifier le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. Ouvrez le fichier **app.css** ? la racine du projet pour sp?cifier les styles personnalis?s du compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

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

## <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier nomm? **my-office-add-in-manifest.xml** pour d?finir les param?tres et les fonctionnalit?s du compl?ment.

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for Project**.

4. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a>D?marrage du serveur de d?veloppement

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>Essayez !

1. Dans Project, cr?ez un projet simple comportant au moins une t?che.

2. Suivez les instructions pour la plateforme que vous utiliserez afin d?ex?cuter votre compl?ment en vue d?en charger une version test dans Project.

    - Windows : [Chargement de version test des compl?ments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online : [Chargement de version test des compl?ments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compl?ments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. Dans Project, s?lectionnez une t?che.

    ![Capture d??cran d?un plan de projet dans Project avec une t?che s?lectionn?e](../images/project_quickstart_addin_1.png)

4. Dans le volet Office, s?lectionnez le bouton **Get Task GUID** pour ?crire le GUID de la t?che dans la zone de texte **Results**.

    ![Capture d??cran d?un plan de projet dans Project avec une t?che s?lectionn?e et le GUID de la t?che ?crit dans la zone de texte dans le volet Office](../images/project_quickstart_addin_2.png)

5. Dans le volet Office, s?lectionnez le bouton **Get Task data** pour ?crire plusieurs propri?t?s de la t?che s?lectionn?e dans la zone de texte **Results**.

    ![Capture d??cran d?un plan de projet dans Project avec une t?che s?lectionn?e et plusieurs propri?t?s de la t?che ?crites dans la zone de texte dans le volet Office](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>?tapes suivantes

F?licitations, vous avez cr?? un compl?ment Project ! Ensuite, d?couvrez les fonctionnalit?s d?un compl?ment Project et explorez des sc?narios plus courants.

> [!div class="nextstepaction"]
> [Compl?ments Project](../project/project-add-ins.md)
