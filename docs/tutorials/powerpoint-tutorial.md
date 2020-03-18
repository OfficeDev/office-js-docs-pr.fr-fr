---
title: Didacticiel sur les compléments PowerPoint
description: Dans ce didacticiel, vous allez créer un complément PowerPoint qui insère une image, insère du texte, obtient les métadonnées des diapositives et navigue entre les diapositives.
ms.date: 12/24/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: a72310c0ab58e544050ec7574841b38560df2fbf
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717396"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a>Didacticiel : Créer un complément de volet de tâches de PowerPoint

Dans ce didacticiel, vous utiliserez Visual Studio pour créer un complément de volet de tâches de PowerPoint qui:

> [!div class="checklist"]
> * Ajout de la photo [Bing](https://www.bing.com) du jour à une diapositive
> * Ajout de texte à une diapositive
> * Get Slide Metadata
> * Naviguer entre les diapositives

## <a name="prerequisites"></a>Conditions requises

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a>Créer votre projet de complément

Procédez comme suit pour créer un projet complément PowerPoint à l’aide de Visual Studio.

1. Choisissez **Créer un nouveau projet**.

2. À l’aide de la zone de recherche, entrez **complément**. Choisissez **Complément web PowerPoint**, puis sélectionnez **Suivant**.

3. Nommez le projet `HelloWorld` et sélectionnez **Créer**.

4. Dans la fenêtre de la boîte de dialogue **Créer un complément Office**, choisissez **Ajouter de nouvelles fonctionnalités à PowerPoint**, puis sélectionnez **Terminer** pour créer le projet.

5. Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.

     ![Didacticiel PowerPoint - Fenêtre de l’explorateur de solutions Visual Studio qui affiche les 2 projets dans la solution HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a>Mise à jour du code 

Modifiez le code de complément comme suit pour créer la structure que vous utiliserez pour implémenter la fonctionnalité de complément dans les étapes suivantes de ce didacticiel.

1. **Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **Home.html**, localisez la balise **div** avec `id="content-main"`, remplacez l’intégralité de la balise **div** avec le balisage suivant et enregistrez le fichier.

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. Ouvrez le fichier **Home.js** à la racine du projet d’application web. Ce fichier spécifie le script pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    (function () {
        "use strict";

        var messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## <a name="insert-an-image"></a>Insérer une image

Procédez comme suit pour ajouter le code qui récupère la photo[Bing](https://www.bing.com) de la journée et insère l’image dans une diapositive.

1. À l’aide de l’explorateur de solutions, ajoutez un nouveau dossier nommé **Controllers** au projet **HelloWorldWeb**.

    ![Didacticiel PowerPoint : Fenêtre de l’explorateur de solutions Visual Studio qui met en évidence le dossier Controllers du projet HelloWorldWeb](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. Cliquez avec le bouton droit de la souris sur le dossier **Controllers**, puis sélectionnez **Ajouter > Nouvel élément généré automatiquement...**.

3. Dans la fenêtre de boîte de dialogue **Ajouter une structure**, sélectionnez **Contrôleur Web API 2 - Vide** et choisissez le bouton **Ajouter**. 

4. Dans la fenêtre de boîte de dialogue **Ajouter un contrôleur**, saisissez **PhotoController** pour le nom du contrôleur, puis sélectionnez le bouton **Ajouter**. Visual Studio crée et ouvre le fichier **PhotoController.cs**.

5. Remplacez tout le contenu du fichier **PhotoController.cs** par le code suivant qui appelle le service Bing pour récupérer la photo du jour en tant que chaîne encodée en base 64. Lorsque vous utilisez l’API JavaScript Office pour insérer une image dans un document, les données de l’image doivent être spécifiées en tant que chaîne encodée en base 64.

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. Dans le fichier **Home.html**, remplacez `TODO1` par le balisage suivant. Ce balisage définit le bouton **Insert Image** (Insérer une image) qui s’affichera dans volet Office du complément.

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. Dans le fichier **Home.js**, remplacez `TODO1` par le code suivant pour attribuer le gestionnaire d’événements pour le bouton **Insert Image** (Insérer une image).

    ```js
    $('#insert-image').click(insertImage);
    ```

8. Dans le fichier **Home.js**, remplacez `TODO2` par le code suivant pour définir la fonction `insertImage`. Cette fonction extrait l’image du service web Bing, puis appelle la fonction `insertImageFromBase64String` pour insérer cette image dans le document.

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. Dans le fichier **Home.js**, remplacez `TODO3` par le code suivant pour définir la fonction `insertImageFromBase64String`. Cette fonction utilise l’API JavaScript Office pour insérer l’image dans le document. Remarque : 

    - l’option `coercionType` spécifiée comme deuxième paramètre de la demande `setSelectedDataAsyc` indique le type de données insérées. 

    - L’objet `asyncResult` encapsule le résultat de la demande `setSelectedDataAsync`, y compris les informations d’état et d’erreur quand la demande a échoué.

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a>Test du complément

1. À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur **F5** ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Dans le volet Office, sélectionnez le bouton **Insert Image** (Insérer une image) permettant d’ajouter la photo Bing du jour sur la diapositive active.

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-insert-image-button.png)

4. Dans Visual Studio, arrêtez le complément en appuyant sur **Shift + F5** ou en choisissant le bouton**Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a>Personnaliser les éléments de l’interface utilisateur (IU)

Procédez comme suit pour ajouter des marques de révision qui personnalisent l’interface utilisateur du volet de tâche.

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

### <a name="test-the-add-in"></a>Test du complément

1. À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur**F5**ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément**Afficher le volet Office** qui apparaît dans le ruban. Le complément est hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Notez que le volet Office contient désormais une section d’en-tête et un titre, et ne contient plus de section de pied de page.

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. Dans Visual Studio, arrêtez le complément en appuyant sur **Shift + F5** ou en choisissant le bouton**Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a>Insérer du texte

Procédez comme suit pour ajouter le code qui insère le texte dans la diapositive titre qui contient l’image[Bing](https://www.bing.com) de la journée.

1. Dans le fichier **Home.html**, remplacez `TODO3` par le balisage suivant. Ce balisage définit le bouton **Insert Text** (Insérer du texte) qui s’affiche dans le volet Office du complément.

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. Dans le fichier **Home.js**, remplacez `TODO4` par le code suivant pour attribuer le gestionnaire d’événements pour le bouton **Insert Text** (Insérer du texte).

    ```js
    $('#insert-text').click(insertText);
    ```

3. Dans le fichier **Home.js**, remplacez `TODO5` par le code suivant pour définir la fonction `insertText`. Cette fonction insère du texte dans la diapositive active.

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a>Test du complément

1. À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur **F5**ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément**Afficher le volet Office** qui apparaît dans le ruban. Le complément est hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Dans le volet Office, sélectionnez le bouton **Insert Image** (Insérer une image) pour ajouter la photo Bing du jour sur la diapositive active et choisissez une mise en page pour la diapositive qui contient une zone de texte pour le titre.

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. Placez votre curseur dans la zone de texte sur la diapositive de titre, dans le volet Office, sélectionnez le bouton **Insert Text** (Insérer du texte) permettant d’ajouter du texte à la diapositive.

    ![Capture d’écran du complément PowerPoint avec le bouton Insert Text (Insérer du texte) sélectionné](../images/powerpoint-tutorial-insert-text.png)


5. Dans Visual Studio, arrêtez le complément en appuyant sur **Shift + F5** ou en choisissant le bouton**Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a>Obtenir les métadonnées des diapositives

Procédez comme suit pour ajouter du code qui extrait les métadonnées pour la diapositive sélectionnée.

1. Dans le fichier **Home.html**, remplacez `TODO4` par le balisage suivant. Ce balisage définit le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive) qui s’affichera dans le volet Office du complément.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. Dans le fichier **Home.js**, remplacez `TODO6` par le code suivant pour attribuer le gestionnaire d’événements pour le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive).

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. Dans le fichier **Home.js**, remplacez `TODO7` par le code suivant pour définir la fonction `getSlideMetadata`. Cette fonction extrait les métadonnées pour la ou les diapositives sélectionnée(s), et les écrit dans une fenêtre de boîte de dialogue contextuelle dans le volet Office du complément.

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### <a name="test-the-add-in"></a>Test du complément

1. À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur **F5**ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément**Afficher le volet Office** qui apparaît dans le ruban. Le complément est hébergé localement sur IIS.

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Dans le volet Office, sélectionnez le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive) pour obtenir les métadonnées pour la diapositive sélectionnée. Les métadonnées de la diapositive sont écrites dans la fenêtre de boîte de dialogue contextuelle en bas du volet Office. Dans ce cas, le tableau `slides` figurant dans les métadonnées JSON contient un objet qui spécifie les éléments `id`, `title` et `index` de la diapositive sélectionnée. Si plusieurs diapositives étaient sélectionnées lorsque vous avez récupéré les métadonnées des diapositives, le tableau `slides` figurant dans les métadonnées JSON contiendrait un objet pour chaque diapositive sélectionnée.

    ![Capture d’écran du complément PowerPoint avec le bouton Get Slide Metadata (Obtenir les métadonnées de la diapositive) mis en évidence](../images/powerpoint-tutorial-get-slide-metadata.png)

4. Dans Visual Studio, arrêtez le complément en appuyant sur **Shift + F5** ou en choisissant le bouton**Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a>Naviguer entre les diapositives

Procédez comme suit pour ajouter le code qui navigue entre les diapositives d’un document.

1. Dans le fichier **Home.html**, remplacez `TODO5` par le balisage suivant. Ce balisage définit les quatre boutons de navigation qui s’afficheront dans le volet Office du complément.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
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

### <a name="test-the-add-in"></a>Test du complément

1. À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur **F5**ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément**Afficher le volet Office** qui apparaît dans le ruban. Le complément est hébergé localement sur IIS.

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

8. Dans Visual Studio, arrêtez le complément en appuyant sur **Shift + F5** ou en choisissant le bouton**Arrêter**. PowerPoint se ferme automatiquement lorsque le complément est arrêté.

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a>Étapes suivantes

Dans ce didacticiel, vous allez créer un complément PowerPoint qui insère une image, insère du texte, obtient les métadonnées des diapositives et navigue entre les diapositives. Pour en savoir plus sur le développement des complément PowerPoint, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Vue d’ensemble des compléments PowerPoint](../powerpoint/powerpoint-add-ins.md)

## <a name="see-also"></a>Voir aussi

* [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
* [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
* [Développement de compléments Office](../develop/develop-overview.md)
