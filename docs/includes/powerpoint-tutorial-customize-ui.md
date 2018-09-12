Dans cette étape du didacticiel, vous allez personnaliser l’interface utilisateur du volet Office.

> [!NOTE]
> Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="customize-the-task-pane-ui"></a>Personnalisation de l’interface utilisateur du volet Office 

1. Dans le fichier **Home.html**, remplacez `TODO2` par le balisage suivant pour ajouter une section d’en-tête et un titre au volet Office. Remarque :

    - Les styles qui commencent par `ms-` sont définis par la [structure Fabric de l’interface utilisateur Office](../design/office-ui-fabric.md), une infrastructure frontale JavaScript pour créer des expériences utilisateur pour Office et Office 365. Le fichier **Home.html** inclut une référence à la feuille de style Fabric.

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. Dans le fichier **Home.html**, recherchez la balise **div** avec `class="footer"` et supprimez toute la balise **div** pour retirer la section de pied de page du volet Office.

## <a name="test-the-add-in"></a>Tester le complément

1. À l’aide de Visual Studio, testez le complément PowerPoint en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Notez que le volet Office contient désormais une section d’en-tête et un titre, et ne contient plus de section de pied de page.

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

