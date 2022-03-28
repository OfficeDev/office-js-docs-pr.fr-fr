---
title: Didacticiel sur les compléments Word
description: Dans ce didacticiel, vous allez cr?er un compl?ment Word qui ins?re (et remplace) des plages de texte, des paragraphes, des images, du code HTML, des tableaux et des contr?les de contenu. Vous découvrirez également comment mettre en forme du texte et comment insérer (et remplacer) du contenu dans les contrôles de contenu.
ms.date: 01/13/2022
ms.prod: word
ms.localizationpriority: high
ms.openlocfilehash: 6fc01db700475d4ff2dda49e471a68d9ae59aa77
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484028"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a>Didacticiel : Créer un complément de volet de tâches Word

Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:

> [!div class="checklist"]
>
> - Insère une plage de texte
> - Formats de texte
> - Remplacer du texte et insérer du texte à divers emplacements
> - Insère des images, du code HTML et des tableaux
> - Crée et met à jour des contrôles de contenu

> [!TIP]
> Si vous avez déjà exécuté le démarrage rapide [Créer votre premier complément du volet des tâches de Word](../quickstarts/word-quickstart.md) et que vous souhaitez utiliser ce projet comme point de départ pour ce didacticiel, accédez directement à la section [Insérer une plage de texte](#insert-a-range-of-text) pour commencer ce didacticiel.

## <a name="prerequisites"></a>Conditions préalables

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Office connecté à un abonnement Microsoft 365 (y compris Office on the web).

    > [!NOTE]
    > Si vous ne disposez pas déjà d’Office, vous pouvez [ rejoindre le programme de développement de Microsoft 365](https://developer.microsoft.com/office/dev-program) pour obtenir un abonnement Microsoft 365 de 90 jours renouvelable gratuit à utiliser pendant son développement.

## <a name="create-your-add-in-project"></a>Créer votre projet de complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project`
- **Sélectionnez un type de script :** `Javascript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Word`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-word.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="insert-a-range-of-text"></a>Insérer une plage de texte

Dans cette étape du tutoriel, vous devez tester par programme que votre complément prend en charge la version actuelle de Word de l’utilisateur, puis insérer un paragraphe dans le document.

### <a name="code-the-add-in"></a>Codage du complément

1. Ouvrez le projet dans votre éditeur de code.

1. Ouvrez le fichier **./src/taskpane/taskpane.html**. Ce fichier contient le balisage HTML pour le volet Office.

1. Recherchez l’élément `<main>` et supprimez toutes les lignes qui apparaissent après la balise `<main>` d’ouverture et avant la balise `</main>` de fermeture.

1. Ajoutez le balisage suivant immédiatement après la balise `<main>` d’ouverture.

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**. Ce fichier contient le code API JavaScript pour Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.

1. Supprimez toutes les références au bouton `run` et à la fonction `run()` en procédant comme suit :

    - Recherchez et supprimez la ligne `document.getElementById("run").onclick = run;`.

    - Recherchez et supprimez la fonction `run()` entière.

1. Dans l’appel de la méthode `Office.onReady`, recherchez la ligne `if (info.host === Office.HostType.Word) {` et ajoutez le code suivant immédiatement après cette ligne. Remarque :

    - La première partie de ce code détermine si la version de Word de l’utilisateur prend en charge une version de Word.js qui inclut toutes les API utilisées à toutes les étapes de ce didacticiel. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur qui appellerait des API non prises en charge. Cela permet à l’utilisateur d’utiliser toujours les parties du complément prises en charge par sa version de Word.
    - La deuxième partie de ce code ajoute un gestionnaire d’événements pour le bouton `insert-paragraph`.

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    ```

1. Ajoutez la fonction suivante à la fin du fichier. Remarque :

   - Votre logique métier Word.js est ajoutée à la fonction qui est transmise à `Word.run`. Cette logique n’est pas exécutée immédiatement. Au lieu de cela, elle est ajoutée à une file d’attente de commandes.

   - La méthode `context.sync` envoie toutes les commandes en file d’attente vers Word pour exécution.

   - L’élément `Word.run` est suivi par un bloc `catch`. Il s’agit d’une meilleure pratique que vous devez toujours suivre.

   [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

    ```js
    async function insertParagraph() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a paragraph into the document.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertParagraph()`, remplacez `TODO1` par le code suivant. Remarque :

   - Le premier paramètre de la méthode `insertParagraph` correspond au texte pour le nouveau paragraphe.

   - Le deuxième paramètre correspond à l’emplacement dans le corps où sera inséré le paragraphe. Les autres options d’insertion de paragraphe, lorsque l’objet parent est le corps, sont « Fin » et « Remplacer ».

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                            "Start");
    ```

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > Si vous testez votre complément sur Mac, exécutez la commande suivante dans le répertoire racine de votre projet avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet. Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Word avec votre complément chargé.

        ```command&nbsp;line
        npm start
        ```

    - Pour tester votre complément dans Word sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre. Remplacez « {url} » par l’URL d’un document Word sur votre OneDrive ou une bibliothèque SharePoint sur laquelle vous avez des autorisations.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. Dans Word, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran de Word avec le bouton Afficher le volet Office en surbrillance.](../images/word-quickstart-addin-2b.png)

1. Dans le volet des tâches, cliquez sur le bouton **Insérer un paragraphe**.

1. Apportez une modification au paragraphe.

1. Sélectionnez à nouveau le bouton **Insérer un paragraphe** . Notez que le nouveau paragraphe apparaît au-dessus du paragraphe précédent, car la méthode `insertParagraph` est insérée au début du corps du document.

    ![Capture d’écran montrant le bouton Insérer un paragraphe dans le complément.](../images/word-tutorial-insert-paragraph-2.png)

## <a name="format-text"></a>Mettre en forme du texte

Dans cette étape du didacticiel, vous devez appliquer un style intégré au texte, appliquer un style personnalisé à texte et modifier la police du texte.

### <a name="apply-a-built-in-style-to-text"></a>Appliquer un style prédéfini au texte

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-paragraph`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="apply-style">Apply Style</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-paragraph`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("apply-style").onclick = applyStyle;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function applyStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to style text.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `applyStyle()` , remplacez `TODO1` par le code suivant. Notez que le code applique un style à un paragraphe, mais que les styles peuvent également être appliqués à des plages de texte.

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ```

### <a name="apply-a-custom-style-to-text"></a>Appliquer un style personnalisé au texte

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `apply-style`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `apply-style`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function applyCustomStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply the custom style.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `applyCustomStyle()` , remplacez `TODO1` par le code suivant. Notez que le code applique un style personnalisé qui n’existe pas encore. Vous allez créer un style portant le nom **MyCustomStyle** à l’étape [Tester le complément](#test-the-add-in-1) .

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ```

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="change-the-font-of-text"></a>Modifier la police du texte

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `apply-custom-style`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="change-font">Change Font</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `apply-custom-style`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("change-font").onclick = changeFont;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function changeFont() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply a different font.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `changeFont()`, remplacez `TODO1` par le code suivant. Notez que le code obtient une référence au deuxième paragraphe à l’aide de la méthode `ParagraphCollection.getFirst` chaînée à la méthode `Paragraph.getNext` .

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ```

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. Si le volet des tâches du complément n’est pas déjà ouvert dans Word, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

1. Assurez-vous qu’il y a au moins trois paragraphes dans le document. Vous pouvez choisir le bouton **Insérer un paragraphe** trois fois. *Vérifiez soigneusement qu’il n’y a pas de paragraphe vide à la fin du document. Si c’est le cas, supprimez-le.*

1. Dans Word, créez un [style personnalisé](https://support.microsoft.com/office/d38d6e47-f6fc-48eb-a607-1eb120dec563) nommé « MyCustomStyle ». Vous pouvez y appliquer la mise en forme que vous souhaitez.

1. Sélectionnez le bouton **Appliquer le style**. Le style prédéfini **Référence intense** est appliqué au premier paragraphe.

1. Sélectionnez le bouton **Appliquer un style personnalisé**. Votre style personnalisé est appliqué au dernier paragraphe. (Si rien ne semble se produire, le dernier paragraphe est peut-être vide. Si c’est le cas, ajoutez-y du texte.)

1. Sélectionnez le bouton **Modifier la police**. La police Courier New, 18 pt, en gras, est appliquée au deuxième paragraphe.

    ![Capture d’écran montrant les résultats de l’application des styles et des polices définis pour les boutons complément Appliquer un style, Appliquer un style personnalisé et Modifier la police.](../images/word-tutorial-apply-styles-and-font-2.png)

## <a name="replace-text-and-insert-text"></a>Remplacer du texte et insérer du texte

Dans cette étape du didacticiel, vous ajouterez du texte dans les plages de texte sélectionnées et en dehors de celles-ci, puis remplacerez le texte de la plage sélectionnée.

### <a name="add-text-inside-a-range"></a>Ajouter du texte dans une plage

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `change-font`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `change-font`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function insertTextIntoRange() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertTextIntoRange()`, remplacez `TODO1` par le code suivant. Remarque :

   - La méthode est destinée à insérer l’abréviation [« (C2R) »] à la fin de la plage dont le texte est « Click-to-Run » (Démarrer en un clic). Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à insérer dans l’objet `Range`.

   - Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Outre « Fin », les autres options possibles sont : « Début », « Avant », « Après » et « Remplacer ».

   - La différence entre « Fin » et « Après » est que « Fin » insère le nouveau texte à la fin de la plage existante, tandis que l’option « Après » crée une plage avec la chaîne et insère la nouvelle plage après la plage existante. De même, « Début » insère le texte au début de la plage existante, tandis que l’option « Avant » insère une nouvelle plage. L’option « Remplacer » remplace le texte de la plage existante par la chaîne dans le premier paramètre.

   - Vous avez vu lors d’une étape précédente du didacticiel que les méthodes insert* de l’objet corps ne disposent pas des options « Avant » et « Après ». Cela est dû au fait que vous ne pouvez pas placer de contenu en dehors du corps du document.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

1. Nous allons ignorer `TODO2` jusqu’à la section suivante. Dans la fonction `insertTextIntoRange()` , remplacez `TODO3` par le code suivant. Ce code est similaire au code que vous avez créé à la première étape du didacticiel, sauf que vous insérez maintenant un nouveau paragraphe à la fin du document plutôt qu’au début. Ce nouveau paragraphe montre que le nouveau texte fait désormais partie de la plage d’origine.

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Ajoutez du code pour récupérer des propriétés de document dans les objets de script du volet Office

Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour *écrire* dans le document Office. Chaque fonction s’est terminée par un appel à la méthode `context.sync()` qui envoie les commandes mises en file d’attente au document à exécuter. Toutefois, le code que vous avez ajouté à la dernière étape appelle la propriété `originalRange.text` , ce qui est une différence significative par rapport aux fonctions précédentes que vous avez écrites, car l’objet `originalRange` est uniquement un objet proxy qui existe dans le script de votre volet Office. Il ne sait pas quel est le texte réel de la plage dans le document, donc sa propriété `text` ne peut pas avoir une valeur réelle. Il est nécessaire d’extraire d’abord la valeur de texte de la plage du document et de l’utiliser pour définir la valeur de `originalRange.text`. Ce n’est qu’à ce point que vous pouvez `originalRange.text` être appelé sans provoquer la levée d’une exception. Ce processus de récupération comporte trois étapes.

1. Mettez en file d’attente une commande pour charger (autrement dit, récupérer) les propriétés que votre code doit lire.

1. Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.

1. Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.

Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.

1. À l’intérieur de la fonction `insertTextIntoRange()`, remplacez `TODO2` par le code suivant.
  
    ```js
    originalRange.load("text");
    await context.sync();

    // TODO4: Move the doc.body.insertParagraph line here.

    // TODO5: Move the final call of context.sync here and ensure
    //        that it does not run until the insertParagraph has
    //        been queued.
    ```

1. Coupez la ligne `doc.body.insertParagraph` et collez-la à la place de `TODO4`.

Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :

```js
async function insertTextIntoRange() {
    await Word.run(async (context) => {

        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        await context.sync();

        doc.body.insertParagraph("Original range: " + originalRange.text, "End");

        await context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
```

### <a name="add-text-between-ranges"></a>Ajouter du texte entre les plages

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-text-into-range`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-text-into-range`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function insertTextBeforeRange() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertTextBeforeRange()`, remplacez `TODO1` par le code suivant. Remarque :

   - La méthode est destinée à ajouter une plage dont le texte est « Office 2019 », avant la plage contenant le texte « Microsoft 365 ». Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à ajouter.

   - Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Pour plus d’informations sur les options d’emplacement, reportez-vous à la discussion précédente sur la fonction `insertTextIntoRange`.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

1. À l’intérieur de la fonction `insertTextBeforeRange()`, remplacez `TODO2` par le code suivant.

     ```js
    originalRange.load("text");
    await context.sync();

    // TODO3: Queue commands to insert the original range as a
    //        paragraph at the end of the document.

    // TODO4: Make a final call of context.sync here and ensure
    //        that it runs after the insertParagraph has been queued.
    ```

1. Remplacez `TODO3` par le code suivant. Ce nouveau paragraphe montre que le nouveau texte n’entre ***pas*** dans la plage sélectionnée d’origine. La plage d’origine contient toujours le texte qu’elle contenait lorsqu’elle avait été sélectionnée uniquement.

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
    ```

1. Remplacez `TODO4` par le code suivant.

    ```js
    await context.sync();
    ```

### <a name="replace-the-text-of-a-range"></a>Remplacer le texte d’une plage

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-text-outside-range`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="replace-text">Change Quantity Term</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-text-outside-range`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("replace-text").onclick = replaceText;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function replaceText() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `replaceText()`, remplacez `TODO1` par le code suivant. Notez que la méthode est destinée à remplacer la chaîne « plusieurs » par la chaîne « beaucoup ». Cela simplifie l’hypothèse que la chaîne est présente et que l’utilisateur l’a sélectionnée.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. Si le volet des tâches du complément n’est pas déjà ouvert dans Word, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

1. Dans le volet Office, cliquez sur le bouton **Insérer un paragraphe** pour vous assurer qu’un paragraphe apparaît au début du document.

1. Dans le document, sélectionnez l’expression « Démarrer en un clic ». *Veillez à ne pas inclure l’espace précédent ou la virgule suivante dans la sélection.*

1. Sélectionnez le bouton **Insérer une abréviation**. L’abréviation « (C2R) » est ajoutée. Notez également qu’en bas du document, un nouveau paragraphe est ajouté avec l’intégralité du texte développé, car la nouvelle chaîne a été ajoutée à la plage existante.

1. Dans le document, sélectionnez l'expression « Microsoft 365 ». *Veillez à ne pas inclure l'espace précédent ou suivant dans la sélection.*

1. Sélectionnez le bouton **Ajouter les informations de version**. L’expression « Office 2019 » est insérée entre « Office 2016 » et « Microsoft 365 ». Notez également qu’en bas du document, un nouveau paragraphe est ajouté. Celui-ci contient uniquement le texte sélectionné à l’origine, car la nouvelle chaîne est devenue une nouvelle plage plutôt que d’être ajoutée à la plage d’origine.

1. Dans le document, sélectionnez le mot « plusieurs ». *Veillez à ne pas inclure l’espace précédent ou suivant dans la sélection.*

1. Sélectionnez le bouton **Modifier la condition de quantité**. Notez que « beaucoup » remplace le texte sélectionné.

    ![Capture d’écran montrant le résultat de la sélection des boutons de complément Insérer une abréviation, Ajouter des informations sur la version et Modifier la condition de quantité.](../images/word-tutorial-text-replace-2.png)

## <a name="insert-images-html-and-tables"></a>Insérer des images, du code HTML et des tableaux

Dans cette étape du didacticiel, vous allez découvrir comment insérer des images, du code HTML et des tableaux dans le document.

### <a name="define-an-image"></a>Définir une image

Procédez comme suit pour définir l’image que vous allez insérer dans le document dans la partie suivante de ce didacticiel.

1. À la racine du projet, créez un fichier nommé **base64Image.js**.

1. Ouvrez le fichier **base64Image.js**, puis ajoutez le code suivant pour spécifier la chaîne codée au format Base64 qui représente une image.

    ```js
    export const base64Image =
        "iVBORw0KGgoAAAANSUhEUgAAAZAAAAEFCAIAAABCdiZrAAAACXBIWXMAAAsSAAALEgHS3X78AAAgAElEQVR42u2dzW9bV3rGn0w5wLBTRpSACAUDmDRowGoj1DdAtBA6suksZmtmV3Qj+i8w3XUB00X3pv8CX68Gswq96aKLhI5bCKiM+gpVphIa1qQBcQbyQB/hTJlpOHUXlyEvD885vLxfvCSfH7KIJVuUrnif+z7nPOd933v37h0IIWQe+BEvASGEgkUIIRQsQggFixBCKFiEEELBIoRQsAghhIJFCCEULEIIBYsQQihYhBBCwSKEULAIIYSCRQghFCxCCAWLEEIoWIQQQsEihCwQCV4CEgDdJvYM9C77f9x8gkyJV4UEznvs6U780rvAfgGdg5EPbr9CyuC1IbSEJGa8KopqBWC/gI7Fa0MoWCROHJZw/lxWdl3isITeBa8QoWCRyOk2JR9sVdF+qvwnnQPsF+SaRSEjFCwSCr0LNCo4rYkfb5s4vj/h33YOcFSWy59VlIsgIRQs4pHTGvYMdJvIjupOx5Ir0Tjtp5K/mTKwXsSLq2hUWG0R93CXkKg9oL0+ldnFpil+yhlicIM06NA2cXgXySyuV7Fe5CUnFCziyQO2qmg8BIDUDWzVkUiPfHY8xOCGT77EWkH84FEZbx4DwOotbJpI5nj5CQWLTOMBj8votuRqBWDP8KJWABIr2KpLwlmHpeHKff4BsmXxFQmhYBGlBxzoy7YlljxOcfFAMottS6JH+4Xh69IhEgoWcesBNdVQozLyd7whrdrGbSYdIqFgkQkecMD4epO9QB4I46v4tmbtGeK3QYdIKFhE7gEHjO/odSzsfRzkS1+5h42q+MGOhf2CuPlIh0goWPSAogcccP2RJHI1riP+kQYdVK9Fh0goWPSAk82a5xCDG4zPJaWTxnvSIVKwKFj0gEq1go8QgxtUQQeNZtEhUrB4FZbaA9pIN+98hhhcatbNpqRoGgRKpdAhUrDIMnpAjVrpJSNApK/uRi7pEClYZIk84KDGGQ+IBhhicMP6HRg1ycedgVI6RELBWl4POFCr8VWkszpe3o76G1aFs9ws+dMhUrDIInvAAeMB0ZBCDG6QBh2kgVI6RAoWWRYPqBEI9+oQEtKgg3sNpUOkYJGF8oADxgOioUauXKIKOkxV99EhUrDIgnhAG+mCUQQhBpeaNb4JgOn3AegQKVhkvj2gjXRLLrIQgxtUQYdpNYsOkYJF5tUDarQg4hCDS1u3VZd83IOw0iFSsMiceUCNWp3WYH0Wx59R6ls9W1c6RAoWmQ8PaCNdz55hiMEN4zsDNhMDpXSIFCwylx5Qo1a9C3yVi69a2ajCWZ43NOkQKVgkph5wwHi+KQ4hBs9SC9+RMTpEChaJlwfUFylWEafP5uMKqIIOPv0sHSIFi8TFAzpLiXxF/KCbdetEGutFUSa6TXQsdKypv42UgZQhfrWOhbO6q8nPqqCD/zU4OkQKFpm9B7SRbrTpQwzJHNaL/VHyiRVF0dfC2xpOzMnKlUgjW0amhGRW/ZM+w5sqzuqTNWtb9nKBZDLoEClYZGYe0EYaENWHGDaquHJv5CPnz/H9BToWkjmsFkTdOX0GS22p1ovYNEdUr9vCeR3dJlIG1gojn2o8RKPiRX+D0iw6RAoWmYEH1HioiQZqq47VW32dalUlfi1fQf7ByEdUQpMpYfOJ46UPcFweKaMSaWyaWL8z/Mibxzgqe3G4CC6pT4dIwSLReUCNWrkJMdjh8sMSuk1d3bReRGb3hy97iS/SEl+5bQ0LqM4B9gvytaptC6kbwz++vD3ZG0r3EBDoWUg6RAoWCd0D9isXReTKTYghZbhdUB/UYlKV2TSHitZtYc9QrqynDGy/GnGg+4XJr779ShJ0gNdAKR3i/PAjXoIZe8BGBS+uhqtWAF4VXUWu3G//ORVqdVRiEumhWgFoVHT7gB1LnFAvVaJxYZJ+qx/XRuo1X0+RFqzPsF/QFZuEgrVcHnDPCGbFylnajN/wAZZvqgpR8IzO275tTvjnwl/4sORC6C9xWJLoYCKNrbpuR3Jazp/jxdUJmksoWIvvAfcLsD4LuLfn5hOJhWlVQ+lyNZDFcUl636GY5/Wpyzo3FRZ+WBeT1JhpGDVlIMMbjYfYM3Ba4zuXgkUPGBD5B5Kl6LaJ4/uh/CCDTvDjW4ROxZm4gj7+dwZLY24067AkF9OtesCaRYdIwaIHDIzMrmSzv2NNTgl4fLlSXw6kjs8pWN+FfHu3n8p/xpSBjWrwL0eHSMGiB/TL+h1JnNJ+xTA6MawXh1ogTWA5S5tvLS8vMVUM6s1j+TKZEASjQ6RgkVl6wH4pcUM+zs8qBq9WyRyMGozP+5J0/nzygrrLSkS4ONPmNg/vyr1npiQG9+kQKVhkBh5woFbSI8EuQwxTkS1j2xoG0zsHeBVcRsl/RNMqyoMOG9WRjAUd4pzD4GhoHjDsMIEqchX48JuUgU1zJN+kSa4D+LnjHfXiqqsa5Oejb8J/fs9TAZjFtiXXvgADpaqXZsqUFRY94NRq1agErFbrRWzVR9Tq9JlOrWy75NncCf982n+o+sYCDJTSIVKw6AGnRhoQbZsBv3S+MlyxAtC7xPF9WMUJDsi5M+gmVCWImpvolorOgXzTMPBAKR0iBWvuPWB4+4CiWj2Rz3MPcFSXHb90NmawbWDLRVZAc2pHZTkF2fWDKugQRqBUCvcQKVj0gI6qRxYQtfvGBIUdvHQ2fmk/VR7fk5Q5jr+2fmfygrpTfM+fu8qa6lEFHcIIlGocolWkQwwcLrr79oBB9YRxg7SDXbDjJISue71LHJWnrno+vRh+BX2Xq2QOO6+Hf3TTXsYl43M3BhVcZFNjEyvIluUNvAgrrIX1gINqRdpvM0C1EhatbBvowaM5neOVe/L2VX176/jip88CUysAhyV5SRheoFRSfV+i8RAvckH+XKyweBW8qNWeEelEP1XkKqgQw3j/T3sxyNv6cSKNm02xA3KrOvLV1gq4Xh1u3vUusWcE7KESK7jZlHvSoDqU+q/4CAUrItomWtUoRvup1KpRCWxb0KiNqFXvcoreWCem/ETh+ILRYJnvJzlxz+7wrt/l9qkuHUIIrMk9bxaZEjIltl2mYMWDjoVWFae1sAouVeQq2LUYZwfRaVG1dR9PnKp802EpxG016TCOgZsOb6tk9RayZVZVFKwZ8cff4b/+Htcq8sd17wInJt5UA17SUqnVWR0vbwf5Qn5KgPO6bo0mU0K2LJetbgtvqjgxQw8uqcbthDH+OrHS/5FV19MuJDXreoSCFQC9C3yxisQK8hVk1dteZ3W8qQY2VFm68OF/emj0JNJ430DKQCKN3gU6FrrNSHf9VaMrfI68F+ynXVKpkhxndRyX0TlQzv4hFKyABWuwMPGROWxiJ6kdmmibaJu+7gTpPRbgDbZsqJa9/T8AMrvIlnWx/m4Tx+XhY4yC5RXGGjzRbeHlbd3ZsWQO+Qp2mth84nFtSBoQtS0M1cobqqCD50BpMovrj/Dpufyk1OBXZueKgyq6KVjEI/bZMf3ef6aErTp2XiOzO8UtIe0gCuCoHMWm5MLWyJfK09HTdihdvwPjc+w0J4wvbJv4KhfF2VIKFnHLm8f4KjfhkF0yh00TN5vYfDJ510wVED0qR7ENv7Sa5SZQmlhB/gF2XsOoTdj+O6tjz8Dh3Tlbaow9XMNy/153rGGpDIJ+Ycv5bm6bcvVR5YaiPFCy8Kze6s+4lj4VpIHS1Vv4sORqa09YrlL5fa5hUbBmLFiDd/am6Soi0LtAqzqyMK9Sq8BDDEQVdMBooDSxgvXihAV14RfqxgBSsChYcREsmyv3lImtcU5raJs4q8sjV/MYYpgLrj9SxlP2C/iuiXxFl1EYL4GPym5/TRQsCla8BKu/3qFNbLl80a9yVKuwUIWzpmKQrnIPBcsrXHQPT+AucXzf70l91lahclT2FV7tNmEV8fI2t24jI8FLEC52Ysv9wpbAtsVLGNNy2+VyFWGFNX+4SWyReYHpKgrWUuAmsUXiDNNVFKwlsxJBLGyRGVh7LlfFAq5hzeTd38LL27oo0ABpnykSIG766pzWYH3GS0XBWvJr7yLg8/1F1J18l4pk1lXuhM1CaQkJPixN/jvXKlGMpVpa8u7CvSkj9CGshIIV92e7tOvxeBXGhGFIrN6Sp0ZPa5Jw1gfsdEzBWmbGb4BuE4d3JbdKtszHe1jllZTjsqTBvJtymFCwFpbxpRM77nAouzE+MnnBAiazK++rYZ9Flw4B4mODgrWkpG5I1nHf1gDFrPa1gveRNmQc+5jnOL2L/pDqzoGkN2mArpChFgrWXD3eS5J38KDJjDTKsMG4aaDlrXTjr1UdJkJPTLpCChYBAEmzSqcHOX8utySZXV65AFBFGezjgULBS1dIwaIflDzehVVeVZHFiIN/VFEGoZtVtyUxbtwrpGDNDb3fheUH26Z4Nq3bkhw5TKT9dtciqihDtynpWN2mK6RgzS/vemH5QemU9kZF0tohX6Er8VteSTmWPQlOZa5w4gwRQsFaZD/Yu5APLOhdyvs6XOfqu+faVhFlOKsrfwXjRRZHzFOwlumeKbkqr2xaVUmOdL3IiEPA5ZXmhPn4b2edy1gUrOVh/O2uaY/Vu2TEITi1eiCPMrRNnD9XC9Yz0Zgnc3SFFKxl9YPd5oT+Su2nkgQjIw7TklhR7ldMbOBzQldIwVpOxu+Z8SWScY7K8iKLEQf3bFTlUYZWdZjXVT4zTLrCGD16eAlm6QfdCJZ9WEdYLbYjDmG3FU/mRqoJD90EV3+Ga//o5aUPS77m2QiFrbQm6l24+ok6B+g2R0pj2xWy9SgFa6HV6o74kO9Ykx/vNsdlyficfGVkanRIgpV/4Euw3v/E4xZBMheYYKn2VZ0HcfS0quK6YaaE4/t8U9MSLlN55X4aRedAXouxVZab54Q0ytBtTnH933KvkIJFwdIEGsaRVjeZEiMOHsurRmWKyTfdlrj1wb1CCtZy+cHT2nSjorotuWbFvMj6w6/xhxN81xL/G/zsvY7ks384wfdBDHBURRmkB3EmukIBHpOaBVzDmlF55Wa5ffyeyZZF4VsrILM79e0XGb/5JX7zS8nHt+r92rDz79gvhPPWVkcZpF0S9cgTpHf51maFtQSCpTqOo0d1WCfPQRUyVFGGs7ouKaq5+IJmJdJYv8PLTMFaDj/ojcZDyd5ZMkd7IqKKMsDHqEcGsihYS+oHT0zvX016v3FQhYBqrV1/EGeCKxw7pkPBomAtGokV8W3dbXq/Z6A4rMNpYE5Wb8mjDPA9SZuucOb3Ey9B6OVVUH5wwFEZW3Xxg5kSTkxfUmjj/MrCdz7+ovpvclxYo2HTVKqVz5xtqyo6zfWil+VIQsGaGz/4xnevBelhHQD5Cl7eDqA88fCpcX6cns0Fv3JPHmUQWrZ7Y/yYDvcKaQkX2Q+6P46j5+uS5IN2xCEO9C7xrTWbC36toiyOpgq+KS25SVfICmtpyqsTM5ivbA/7HN8Iy1emjqQKOGu0lIHrj+SfEhD+5mFJ0t85AlQDJrrNwA6Kt01xuZCukIK1sILlIS+qolGRLJDZEQc/N6dmxqfmU85dufbTANbpPKCa3wXfa+3Co6JjIWX4coWzWt2jJSRT+EGftc/4nSNdlMmWo86R5ivDg3XdlryBVwR8ZCrVIdiTACdjrnBaJx7g24CCRcIqrwKvO1pVifNKpCPtoZwyRlrQfD0jM6iJMgQuoEyQUrAWX7B6F8ELVu8S38jMTqYUXS8BZ4ag8VBnGyP7NgQb6z/qMX7ZhV/lepGnoyhYMeP/vouRHxzw5rG80V0008CcZrBzEORS0VSoogxQDBz0D6fpULAWSrAi8IPDukYmE2uF0LfbBTPooQVCIGiiDG0zrEbG7ac8pkPBWiCEwEG3GeLOd/up3IiFXWQ5Xdjx/ZntfKmiDEC4FR9dIQVrQUhmxQXgsLf5pXem0JE9PDN4/jyAELnnS62JMoTa8P7EpCukYC0EH4QZv5JiH9YZJ6SIg9MM9i5nZgY1VWQgB3EmXnNh9ZCCRcGaSz4cvYE7VhQjoaSHdUKKODjNYIDzuKZl9ZZSI76pRJF1oiukYC2CH3TGoBHccRw99mGdcQKPODjN4Omz2YTabVRa3G3izeMovoHxc+wssihYc+8H30Z1Szcq8tBmgKvv8TGDmV3xweC8DtEwPk2HgkXBmm8/eFoLd+lXuH+kCzcBRhycZtAqzibUDiCxoiyvzuqRjuQQyuf1Ilu/UrDm2Q9G7Jikh3WCKrKcZvDN41BC7X/+NzBq+Nk3yurJZnx6UPTllap8/oBFFgVrfv1gxILVu5QfnUvmcOWe3y8+CBB0DuRHgvyI1F//Cp9+i7/6Bdbv4E/zuv5/yayyH3QYB3EmVrXCr/jDEu8DCtZ8+sG2OYNz+e2n8m27a76ngQ3+eYDtrlZv9UXqp3+BRMrVP9FUi1/PQiwEwUoZdIUULPrBaZAeoAtqUEXj4SzbOWmiDG0zuuVC4bcsyDddIQVrDhCO43iblhrMLfRMmSP1+fCP4ITz//4WHUuZ7dpQJ0VndfR6vHkDXSEFa/4E68Sc5Tejuns/Mn3dmVY4tUOvg9//J379C/zbTdQ/wN7HcsHSRBla1dmUV3SFFKy5JHVD7HAS9nEcPefP5YZ0rTDd8BtBBIMKtf/oJwDwP/+N869w/Hf44n3861/iP/4WFy+U/0QTZfB/EGe9qOyo5bKkFa4MXWE4sKd7OOVVtxnFcRw9x2X5cs+miRdXXX2Fb62RwRMB5hga/4Df/2o6+dNEGfwfxLle7ddEnqOwp7WRY9gfliJK27PCIh4f0YJDmTmqwzruIw69C5zVh/8FyG//aTq10nRl8H8QJ1/pq1VmVzKIyCXCpaYrpGDNkx98W4vFN3ZUlucPrlXm7JhueE2vEukRKfS8kdo5EDdPPWsfoWBF6gfP6gEvAKcM5Cv9/zIl5a0rKZEu5bVeUBGHaFi9pbz5/R/E2aiOaHcy611oTkwKVti89+7dO14Fd49QC3sfyz+183qkwjosBXacba2AfEVcJrdlSHUKR9SmFdxsyjXuRW6WO2vu+eRL5USc/YKvaHvKwPYriZV+kfPy1ZJZ7Iz63D1DuZT5c953rLBi4gcDyYsmc9g08cmXkk29xAryD3CzqbyNBXVTzbnyE3GIrnrdVf6YpzW/B3Gc247dVl++PRdZ3Za40qf5OrM6N07Boh8U7yKfO1a2VO28njCeM7GCT750dWupDuv4iThEQ2JFZ119TsRZL478+F+Xhsthnv2ysPSu6TbzLYc/U7BmgvCm9Bm/ShnYtiRS1TlA4yEaD3H+fEQQN5+46imq2q3fqMb62mbLyvld/g/iOM8k2mcDBl/Tc5ElFNfJXHQDIilYxIVa3Rm5o3wex0kZ2KqL+3ftp3hxFXsGGhU0Ktgv4Is0Xt4eytaVe5MrAlXT95Qx9Zj1yNBEGXoXk+c5pwydZR5EGWzXPCjWfBZZvUvxicWldwrWbHjXm1xe+Vy92jRH1KpzgL2P5U3Tz+ojp2TyD5SVyADV9r+wTRYfNFGGVnWC706kYdTwyZfYqktkS4gytKrDKzxw9EEVWexBSsGaDb3fTRYsP3lRofl65wD7BV1fBGFH302RJbWrwt0bEzRRBjcHca79UECt3pLIllOju60RKXd+cW9F1umzkQV1ukIKVoz8oLME8Hkcx6l9vUvsFyZvJDnv29XC5JdQFVlOfxSf8krFUXlCeZXMiWLnlC3BBY+30BqUb56LrBO6QgpWHAUr0OV2Z49NVUJdoGMNb103iqNq+o7wx0RPV2yqowzd5uSMW7eJPUOymDiQLWc1NL6057/Icr9XSChY8ypYmnUQvWYNcBPLUk3WEfb4Z0ggUYZuE1YR1meSWmxgBp1r7SrF8VZkdQ5Glh2TubjHRyhYS+cHO5bfXXan9LhPFTrvBDfHiVWHdRCbiIMmynBWn24T9rSGr3LKo9HfXygX9Z11nLciS7jIbOlHwYpXeeW/PcP3DpHSz4xRlVQu+x84N8WcxCHikFjR7QB4OOdsByBe3pYsLyaz2H6FTVOuj4PX8lZkveVeIQUrzoI10cQl0hNaxDkrLDfbdon0yMKT+0Mqvcv4Rhw2qsqqx89BnLM69gx5CZzZxc5ryev6LLKEGauJdGCjISlYxK8fnHgcZ72Im01dh1+MtsfL7E7OVW1UR/bLT8wpvn/VYZ3ZRhxSN3S1jM+DOGuF4b6EcFoAwJV7uNkUk1+DqtlbkSUU3SyyKFhzU14Zn/crF826eO9iZP9r09S1kcmWR+zb6bOpl/xVh3VmGHHQ7FT6b9k+qJJ6l3hVxJ4h7jYOjpQPtKljDWs6D0UWE6QUrFiQWBl53gpCI7d7Pyyg6B/UDUer39Vb2KpLNCuRxkYV1x+NfHEPjX1Vh3Uwo4jD+h2lmvufiOM85m235ek2cVjCy9uizUysYPMJdn6QLT8rWcI0HbpCCtZ8lFdOd5C6oSuy7LvIaZGcD/y1AjIlbFsjDY57l97HmqpM1kwiDvryymcDDLuNcrclbpKe1bFfwOFd8esns9h80k9s+SmyGMgKGjbwc81ZvT+Rwfh85J3npodcIo2bzb4rPH+O/cIEQRQOFWqe4frjOxPZfCIvHAY/bDTkHyjlwE6BBjVAO5nTLd7lH8i+gdbQIx/endp6f3o+LJN7F/hitf//mq6EhBVWkH7QqVbdpqutK2d4WjO7eFCyfZVD4+GEgz7+1QrqoMBaIbqIw8QoQ1BqBXXyw3adL65KfpvOFT2fK1l0hRSsOfCD475m05zwdLXvnz0DL66i8VByx3YOsGcEMDJeOPo7UvVENahCE2VwcxAnQLpN7Bfw8rZygd/DShb3CilYMRKsN67Xp3sXw/Upu1mopn2KfXzXqGHnNfIPROGwTWVQM01VveGTuSgiDvoog+cpgT69/4scju8HU9kJx3TWi3M2ryhmcA1rmvexVcSnjntbM5ZCxaY5YrXsjaSOhY6FRBopA8kcUoauIUnjod8tM0kxpVhC6l0o85ZBoVnKiXgdTeJV09iojvy+vM2nEC6vPaOEa1gUrNAFq22OpNWPyl5GeAqa5Z7z52hUAh5oOkAY/DOgbeLwbmjl6h0Yak/tcyJOYDWggY1qf9vUw6I7xqbpnNZgfUbBoiWM3A96a89wWJrabpw+w8vb2C+EpVZQr75nSiFGHDRRhrYZC7Wy6+j9AqzPvKRzB3WZc7WRrpAVVhRc/AvSPxOfk37sxnoRawUkc0ikJR6w28J5HWd1nNYiGgm1/Up+cigka3blnq4/xLzMTPT2wx6WkCmxwqJghcnvj/DTDXElItgVk/cNAPjWms3QOjtbr6oKA/5h1eNdAbSqOL6/UG+exMrI6udpDYk0BYuCFSZ//B3+5M/6/9+7wFe5IPNBMUG1sBJsehPA9Ue6iTgLeW2FvHHHcttEiDjgGpZrBmqFIKalxhPVYZ1gIw6a+V0I4iBOPBEie1QrCtbM3nwLQ+dAua6cLQfWxeEjU/mpbhONh4t5bdtPOZ6egjULuk1f01JjjqrpeyLtfYC7k9VburWbwCNmfM5RsFheLbQcqyfrCJMTvaFpu9qxIj2IEz0nJu8eClb0tf2iv+1Uh3Xgu1XWlXu6TqpH5QW/sOfPAztQRcEiruhYvqalzgW9S3yjsGZrBe/9BhIruKZ2fGf1uCRFWZ5TsFjVzxlvHitrAc9FluawN3y3bGd5TsEiEt4uzRNStf6dzMkb3enRRxna5uLXrf0K/SCApkAULOK2nl+k8yITaoGnyqOL2fLUp+E+Mr2II4t0QsHyJVhLhUpH7L4r7pkYZViex8BSFekULApWpGgm60wVcdCom7N59JLQbXHp3TMJXgK3vOvBqKF3gY6FbhPdJr5rLn5p8HVppJeTk+tVV10c9ONjF/UgzshNtoKUgR+nkTKGbRqJJ3j42f8Ds4luEx2rr2XfX6BjLdRNqJqsA8AqTgj967sydJt4cXWh3gypG8M2DKsFAGzJQMGaE2wzdV7v/3/vYl43wpJZbFty0ZmoOJr5XQiha02U1+QnOSRz/ZbWdmsgTWiDULDmkt5Fv93VfPlKje40KsrjykJr4HFBn23Lds9ujoaOgkVfGWtfqXF2mvZVQgcogZi0bKebo2CRBfSVmo7G0gahmv6lsy2v6OYoWMuL7ewiftPPyleqJutA1oJd1SFe9fcXz83ZD5vvmlPPXiUUrBBpm8Pooz1gZmAr7LtlYXylZiqXUDFldnVtZAIfHTZbN6e67IkVZMvIllm+UbDiR6uKRkWuDs5HfTI39CPz6Cs10/QGa1L6KIOf4ayzdXNTFbaZXWxUKVUUrBhjh7bdJyHt289pW+LvKzUrU4OIgz7KoNlVjJub8ybxmV3kK9xJpGDNj2wdlX3Fi2LuKzV7f0dlvK3pogzjW4rxdHOef3H5CvcWKVhzSLeJ43KQrd/j4yuTOeUqsl21ae7YjoXT2tyUk1N51Y9MShUFa845q6NRCTdtNFtfGc9rjgiDIMks8hXuA1KwFojTGo7LUcfZZ+srI3Nz3/3g6aKP2nITkIK1yLRNHJVnHF6fua/06eZsVYrDYaYr93CtQqmiYC00024jRkZMfKUtSQM3B8RxLAU3ASlYSydb31Tw5vEcfKsh+cqZuznPV2OjyhHzFKylpNtEozKXzVXc+8p4ujkPpG7gepWbgBSspSeCbcRoGA+LzkX3GDdmmZuAsXpc8hLMkrUC1uo4q+Pr0nINYpiLQjJb1kX2ySzgEIp4yNZOE5tPkMzyYsSlYLzZpFpRsIiaTAnbFvIPph75R4L8Lexi5/WEIdWEgkUAIJFGvoKbTS+jlYlPVm9h5zU2TUYWKFhketnaeY3MLi9GRFL1yZfYqlOqKFjEK8kcNk1sv+qHoUgoFzmLzSfYqjOyQMEiQZAysFXHJ19OMWaZuCpjV3D9EXbYv5iCRQJnrYBti9uIgUmVvYzBIcUAAAIqSURBVAmYLfNiULBIaGRK2GlyG9HfNdzFtsVNQAoWiYrBNiJlayq4CUjBIjMyNWnkK9i2uI3oVqq4CUjBIjPG3kbcec1tRPUlysL4nJuAFCwSJ9mytxEpWyNF6Ao2n2CnqZyXQShYZGasFbBV5zZiX6rsTUDmFShYJNbY24jXHy3venxmt39omZuAFCwyH2TLy7iNuH6nvwlIqaJgkXmzRcu0jWhvAho1bgJSsMg8M9hGXL+zoD9gtp9X4CYgBYssjmwZtUXbRrQPLe80KVUULLKI2NuIxudzv41obwJuW9wEpGCRRWe92O/FPKfr8VfucROQgkWWjExp/rYR7c7FG1VKFQWLLB+DXszx30a0NwF5aJlQsChb/W3EeMpW6gY3AQkFi4xipx9itY1obwJuW5QqIj5keQkIEJuRrhxfSlhhkSlka4YjXTm+lFCwyNREP9KV40sJBYv4sGY/bCNeuRfuC63ewvYrbgISChYJQrY2qmFtIw46F6cMXmlCwSIBEfhIV44vJRQsEi6BjHTl+FJCwSLR4XmkK8eXEgoWmQ3TjnTl+FJCwSIzZjDSVQPHl5JAee/du3e8CsQX3Sa6Y730pB8khIJFCKElJIQQChYhhFCwCCEULEIIoWARQggFixBCwSKEEAoWIYRQsAghFCxCCKFgEUIIBYsQQsEihBAKFiGEULAIIRQsQgihYBFCCAWLEELBIoQQChYhhILFS0AIoWARQkjA/D87uqZQTj7xTgAAAABJRU5ErkJggg==";
    ```

### <a name="insert-an-image"></a>Insérer une image

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `replace-text`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Recherchez l’appel de méthode `Office.onReady` en haut du fichier, puis ajoutez le code suivant immédiatement avant cette ligne. Ce code importe la variable que vous avez définie précédemment dans le fichier **./base64Image.js**.

    ```js
    import { base64Image } from "../../base64Image";
    ```

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `replace-text`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("insert-image").onclick = insertImage;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function insertImage() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert an image.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertImage()`, remplacez `TODO1` par le code suivant. Notez que cette ligne insère l’image encodée en base 64 à la fin du document. (L’objet `Paragraph` a également une méthode `insertInlinePictureFromBase64` et d’autres méthodes `insert*` . Pour obtenir un exemple, consultez la section insertHTML suivante.)

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a>Insérer du code HTML

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-image`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="insert-html">Insert HTML</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-image`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("insert-html").onclick = insertHTML;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function insertHTML() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a string of HTML.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertHTML()`, remplacez `TODO1` par le code suivant. Remarque :

   - La première ligne ajoute un paragraphe vide à la fin du document.

   - La deuxième ligne insère une chaîne de code HTML à la fin du paragraphe. Plus précisément, deux paragraphes : un paragraphe avec la police Verdana, et l’autre avec le style par défaut du document Word. (Comme pour la méthode `insertImage` précédente, l’objet `context.document.body` contient également les méthodes `insert*`.)

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a>Insérer une forme

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-html`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="insert-table">Insert Table</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-html`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("insert-table").onclick = insertTable;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function insertTable() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `insertTable()` , remplacez `TODO1` par le code suivant. Notez que cette ligne utilise la méthode `ParagraphCollection.getFirst` pour obtenir une référence au premier paragraphe, puis utilise la méthode `Paragraph.getNext` pour obtenir une référence au deuxième paragraphe.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

1. Dans la fonction `insertTable()`, remplacez `TODO2` par le code suivant. Remarque :

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

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. Si le volet des tâches du complément n’est pas déjà ouvert dans Word, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

1. Dans le volet des tâches, cliquez sur le bouton **Insérer un paragraphe** au moins trois fois pour vous assurer qu’il existe quelques paragraphes dans le document.

1. Sélectionnez le bouton **Insérer une image** et vous remarquerez qu’une image est insérée à la fin du document.

1. Sélectionnez le bouton **Insérer du code HTML**, puis notez que deux paragraphes sont insérés à la fin du document, et que le premier est affiché dans la police Verdana.

1. Sélectionnez le bouton **Insérer un tableau** et notez qu’un tableau est inséré après le deuxième paragraphe.

    ![Capture d’écran illustrant le résultat de la sélection des boutons de complément Insérer une image, Insérer du code HTML et Insérer un tableau.](../images/word-tutorial-insert-image-html-table-2.png)

## <a name="create-and-update-content-controls"></a>Créer et mettre à jour des contrôles de contenu

Dans cette étape du didacticiel, vous découvrirez comment créer des contrôles de contenu de texte enrichi dans le document, puis comment insérer et remplacer du contenu dans les contrôles.

> [!NOTE]
> Plusieurs types de contrôles de contenu peuvent être ajoutés à un document Word via l’interface utilisateur. Toutefois, actuellement, seuls les contrôles de contenu de texte enrichi sont pris en charge par Word.js.
>
> Avant de commencer cette étape du didacticiel, nous vous recommandons de créer et de manipuler des contrôles de contenu de texte enrichi via l’interface utilisateur Word afin de vous familiariser avec les contrôles et leurs propriétés. Pour plus d’informations, reportez-vous à l’article [Créer des formulaires à remplir ou imprimer dans Word](https://support.microsoft.com/office/040c5cc1-e309-445b-94ac-542f732c8c8b).

### <a name="create-a-content-control"></a>Créer un contrôle de contenu

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `insert-table`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="create-content-control">Create Content Control</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `insert-table`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("create-content-control").onclick = createContentControl;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function createContentControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to create a content control.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `createContentControl()`, remplacez `TODO1` par le code suivant. Remarque :

   - Ce code est destiné à intégrer l’expression « Microsoft 365 » dans un contrôle de contenu. Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - La propriété `ContentControl.title` indique le titre visible du contrôle de contenu.

   - La propriété `ContentControl.tag` indique une balise qui peut être utilisée pour obtenir une référence à un contrôle de contenu à l’aide de la méthode `ContentControlCollection.getByTag`, que vous utiliserez dans une fonction ultérieure.

   - La propriété `ContentControl.appearance` indique l’apparence visuelle du contrôle. Utiliser la valeur « Tags » (Balises) signifie que le contrôle est intégré entre des balises de début et de fin, et que la balise de début portera le titre du contrôle de contenu. Les autres valeurs possibles sont « BoundingBox » (Cadre englobant) et « None » (Aucun).

   - La propriété `ContentControl.color` spécifie la couleur des balises ou la bordure du cadre englobant.

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a>Remplacer le contenu du contrôle de contenu

1. Ouvrez le fichier **./src/taskpane/taskpane.html**.

1. Recherchez l’élément `<button>` pour le bouton `create-content-control`, et ajoutez le balisage suivant après cette ligne.

    ```html
    <button class="ms-Button" id="replace-content-in-control">Rename Service</button><br/><br/>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js**.

1. Au cours de l’appel de méthode `Office.onReady`, recherchez la ligne qui attribue un gestionnaire de clic au bouton `create-content-control`, puis ajoutez le code suivant après cette ligne.

    ```js
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
    ```

1. Ajoutez la fonction suivante à la fin du fichier.

    ```js
    async function replaceContentInControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dans la fonction `replaceContentInControl()`, remplacez `TODO1` par le code suivant. Remarque :

    - La méthode `ContentControlCollection.getByTag` renvoie un élément `ContentControlCollection` comprenant tous les contrôles de contenu de la balise spécifiée. Nous utilisons `getFirst` pour obtenir une référence pour le contrôle souhaité.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

1. Vérifiez que vous avez enregistré toutes les modifications que vous avez apportées au projet.

### <a name="test-the-add-in"></a>Test du complément

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. Si le volet des tâches du complément n’est pas déjà ouvert dans Word, sélectionnez l’onglet **Accueil**, puis cliquez sur le bouton **Afficher le volet de tâches** du ruban pour l’ouvrir.

1. Dans le volet des tâches, cliquez sur le bouton **Insérer un paragraphe** pour vous assurer qu’il existe un paragraphe contenant « Microsoft 365 » en haut du document.

1. Dans le document, sélectionnez le texte « Microsoft 365 », puis choisissez le bouton **Créer un contrôle de contenu**. Notez que l’expression est encapsulée dans des balises étiquetées « Nom du service ».

1. Sélectionnez le bouton **Renommer le service** et notez que le texte du contrôle de contenu devient « Fabrikam Online Productivity Suite ».

    ![Capture d’écran illustrant le résultat de la sélection des boutons de complément Créer un contrôle de contenu et Renommer le service.](../images/word-tutorial-content-control-2.png)

## <a name="next-steps"></a>Étapes suivantes

Dans ce didacticiel, vous avez créé un Word tâche volet complément qui insère et remplace le texte, images et autres content dans un document Word. Pour en savoir plus sur la création de compléments Word, passez à l’article suivant.

> [!div class="nextstepaction"]
> [Présentation des compléments Word](../word/word-add-ins-programming-overview.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
