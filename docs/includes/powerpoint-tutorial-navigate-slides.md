Dans cette étape du didacticiel, vous allez parcourir les diapositives d’un document.

> [!NOTE]
> Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.

## <a name="navigate-between-slides-of-the-document"></a>Naviguer entre les diapositives du document

1. Dans le fichier **Home.html**, remplacez `TODO5` par le balisage suivant. Ce balisage définit les quatre boutons de navigation qui s’afficheront dans le volet Office du complément.

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-first-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to First Slide</span>
        <span class="ms-Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-next-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Next Slide</span>
        <span class="ms-Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-previous-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Previous Slide</span>
        <span class="ms-Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-last-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Last Slide</span>
        <span class="ms-Button-description">Go to the last slide.</span>
    </button>
    ```

2. Dans le fichier **Home.js**, remplacez `TODO8` par le code suivant pour affecter les gestionnaires d’événements pour les quatre boutons de navigation.

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. Dans le fichier **Home.js**, remplacez `TODO9` par le code suivant pour définir les fonctions de navigation. Chacune de ces fonctions utilise la fonction `goToByIdAsync` pour sélectionner une diapositive en fonction de sa position dans le document (première, dernière, précédente, suivante).

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>Tester le complément

1. À l’aide de Visual Studio, testez le complément en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)


3. Utilisez le bouton **Nouvelle diapositive** dans le ruban de l’onglet **Accueil** pour ajouter deux nouvelles diapositives au document. 

4. Dans le volet Office, sélectionnez le bouton **Go to First Slide** (Aller à la première diapositive). La première diapositive du document est sélectionnée et affichée.

    ![Capture d’écran du complément PowerPoint avec le bouton Go to First Slide (Aller à la première diapositive) mis en évidence](../images/powerpoint-tutorial-go-to-first-slide.png)

5. Dans le volet Office, sélectionnez le bouton **Go to Next Slide** (Aller à la diapositive suivante). La diapositive suivante du document est sélectionnée et affichée.

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Next Slide (Aller à la diapositive suivante) mis en évidence](../images/powerpoint-tutorial-go-to-next-slide.png)

6. Dans le volet Office, sélectionnez le bouton **Go to Previous Slide** (Aller à la diapositive précédente). La diapositive précédente du document est sélectionnée et affichée.

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Previous Slide (Aller à la diapositive précédente) mis en évidence](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. Dans le volet Office, sélectionnez le bouton **Go to Last Slide** (Aller à la dernière diapositive). La dernière diapositive du document est sélectionnée et affichée.

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Last Slide (Aller à la dernière diapositive) mis en évidence](../images/powerpoint-tutorial-go-to-last-slide.png)

8. Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)
