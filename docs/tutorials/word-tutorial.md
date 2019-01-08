---
title: Didacticiel sur les compléments Word
description: Dans ce didacticiel, vous allez cr?er un compl?ment Word qui ins?re (et remplace) des plages de texte, des paragraphes, des images, du code HTML, des tableaux et des contr?les de contenu. Vous découvrirez également comment mettre en forme du texte et comment insérer (et remplacer) du contenu dans les contrôles de contenu.
ms.date: 12/31/2018
ms.topic: tutorial
ms.openlocfilehash: d1d278d1acd9e8a1377773b90ae9d528af69b93c
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724947"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a>Didacticiel : Créer un complément de volet de tâches Word

Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:

> [!div class="checklist"]
> * Insère une plage de texte
> * Formats de texte
> * Remplacer du texte et insérer du texte à divers emplacements
> * Insère des images, du code HTML et des tableaux
> * Crée et met à jour des contrôles de contenu 

## <a name="prerequisites"></a>Conditions requises

Pour utiliser ce didacticiel, les logiciels suivants doivent être installés. 

- Word 2016, version 1711 (Démarrer en un clic version 8730.1000) ou version ultérieure. Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).

- [Node](https://nodejs.org/en/) 

- [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

## <a name="create-your-add-in-project"></a>Créer votre projet de complément

Procédez comme suit pour créer le projet de complément Word que vous souhaitez utiliser comme base pour ce didacticiel.

1. Clonez le référentiel GitHub du [didacticiel sur les compléments Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).

2. Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

3. Exécutez la commande `npm install` pour installer les outils et les bibliothèques répertoriées dans le fichier package.json. 

4. Effectuez les étapes décrites dans la rubrique relative à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.

## <a name="insert-a-range-of-text"></a>Insérer une plage de texte

Dans cette étape du tutoriel, vous devez tester par programme que votre complément prend en charge la version actuelle de Word de l’utilisateur, puis insérer un paragraphe dans le document.

### <a name="code-the-add-in"></a>Codage du complément

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. Remplacez `TODO1` par le codage suivant :

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

6. Remplacez `TODO2` par le code suivant :

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Remplacez `TODO3` par le code suivant. Remarque:

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

8. Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :

   - Le premier paramètre de la méthode `insertParagraph` correspond au texte pour le nouveau paragraphe.

   - Le deuxième paramètre correspond à l’emplacement dans le corps où sera inséré le paragraphe. Les autres options d’insertion de paragraphe, lorsque l’objet parent est le corps, sont « Fin » et « Remplacer ».

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Chargez une version test du complément en utilisant l’une des méthodes suivantes :

    - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

    - Word Online : [Chargement d’une version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)

    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

5. Dans le menu **Accueil** de Word, sélectionnez **Afficher le volet des tâches**.

6. Dans le volet Office, sélectionnez **Insérer un paragraphe**.

7. Apportez une modification au paragraphe.

8. Sélectionnez à nouveau **Insérer un paragraphe**. Notez que le nouveau paragraphe se trouve au-dessus du précédent, car la méthode `insertParagraph` effectue l’insertion au « début » du corps du document.

    ![Didacticiel Word- Insérer un paragraphe](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a>Mettre en forme du texte

Dans cette étape du didacticiel, vous devez appliquer un style intégré au texte, appliquer un style personnalisé à texte et modifier la police du texte.

### <a name="apply-a-built-in-style-to-text"></a>Appliquer un style prédéfini au texte

1. Ouvrez le projet dans votre éditeur de code. 

2. Ouvrez le fichier index.html.

3. Juste en dessous de la balise `div` qui contient le bouton `insert-paragraph`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-paragraph`, ajoutez le code suivant :

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. Ajoutez la fonction suivante juste après la fonction `insertParagraph` :

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. Remplacez `TODO1` par le code suivant. Le code applique un style à un paragraphe, mais les styles peuvent également être appliqués aux plages de texte.

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a>Appliquer un style personnalisé au texte

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `apply-style`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-style`, ajoutez le code suivant :

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. Sous la fonction `applyStyle`, ajoutez la fonction suivante :

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

6. Remplacez `TODO1` par le code suivant. Le code applique un style personnalisé qui n’existe pas encore. Vous allez créer un style nommé **MyCustomStyle** lors de l’étape [Test du complément](#test-the-add-in).

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a>Modifier la police du texte

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `apply-custom-style`, ajoutez le balisage suivant :

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-custom-style`, ajoutez le code suivant :

    ```js
    $('#change-font').click(changeFont);
    ```

5. Sous la fonction `applyCustomStyle`, ajoutez la fonction suivante :

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

6. Remplacez `TODO1` par le code suivant. Le code obtient une référence au deuxième paragraphe en utilisant la méthode `ParagraphCollection.getFirst` chaînée à la méthode `Paragraph.getNext`.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.   

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Assurez-vous qu’il existe au moins trois paragraphes dans le document. Vous pouvez sélectionner trois fois l’option **Insérer un paragraphe**. *Vérifiez attentivement qu’aucun paragraphe vide n’apparaît à la fin du document. S’il y en a un, supprimez-le.*

6. Dans Word, créez un style personnalisé nommé « MyCustomStyle ». Vous pouvez y appliquer la mise en forme que vous souhaitez.

7. Sélectionnez le bouton **Appliquer le style**. Le style prédéfini **Référence intense** est appliqué au premier paragraphe.

8. Sélectionnez le bouton **Appliquer un style personnalisé**. Votre style personnalisé est appliqué au dernier paragraphe. (Si rien ne semble se produire, le dernier paragraphe est peut-être vide. Si c’est le cas, ajoutez-y du texte.)

9. Sélectionnez le bouton **Modifier la police**. La police Courier New, 18 pt, en gras, est appliquée au deuxième paragraphe.

    ![Didacticiel Word- Appliquer des styles et une police](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a>Remplacer du texte et insérer du texte

Dans cette étape du didacticiel, vous ajouterez du texte dans les plages de texte sélectionnées et en dehors de celles-ci, puis remplacerez le texte de la plage sélectionnée.

### <a name="add-text-inside-a-range"></a>Ajouter du texte dans une plage

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. En dessous de la balise `div` qui contient le bouton `change-font`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Sous la ligne qui attribue un gestionnaire de clics au bouton `change-font`, ajoutez le code suivant :

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. Sous la fonction `changeFont`, ajoutez la fonction suivante :

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

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

7. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - La méthode est destinée à insérer l’abréviation [« (C2R) »] à la fin de la plage dont le texte est « Click-to-Run » (Démarrer en un clic). Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à insérer dans l’objet `Range`.

   - Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Outre « Fin », les autres options possibles sont : « Début », « Avant », « Après » et « Remplacer ». 

   - La différence entre « Fin » et « Après » est que « Fin » insère le nouveau texte à la fin de la plage existante, tandis que l’option « Après » crée une plage avec la chaîne et insère la nouvelle plage après la plage existante. De même, « Début » insère le texte au début de la plage existante, tandis que l’option « Avant » insère une nouvelle plage. L’option « Remplacer » remplace le texte de la plage existante par la chaîne dans le premier paramètre.

   - Vous avez vu lors d’une étape précédente du didacticiel que les méthodes insert* de l’objet corps ne disposent pas des options « Avant » et « Après ». Cela est dû au fait que vous ne pouvez pas placer de contenu en dehors du corps du document.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. Nous ignorerons `TODO2` jusqu’à la section suivante. Remplacez `TODO3` par le code suivant. Ce code est similaire au code que vous avez créé lors de la première phase du didacticiel, sauf que, maintenant, vous insérez un nouveau paragraphe à la fin du document plutôt qu’au début. Ce nouveau paragraphe montre que le nouveau texte fait désormais partie de la plage d’origine.

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Ajouter du code pour récupérer des propriétés de document dans les objets de script du volet Office

Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour écrire (*write*) dans le document Office. Chaque fonction se terminait par un appel de la méthode `context.sync()` qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées. Cependant, le code que vous avez ajouté dans la dernière étape appelle la propriété `originalRange.text` et c’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `originalRange` est uniquement un objet de proxy qui existe dans le script de votre volet Office. Il ne connaît pas le texte réel de la plage dans le document, donc sa propriété `text` ne peut pas contenir de valeur réelle. Il est nécessaire de récupérer d’abord la valeur de texte de la plage à partir du document, puis de l’utiliser pour définir la valeur de `originalRange.text`. Seulement ensuite, la propriété `originalRange.text` peut être appelée sans générer d’exception. Ce processus de récupération comporte trois étapes :

   1. Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.

   2. Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.

   3. Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.

Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.

1. Remplacez `TODO2` par le code suivant.
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.

            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

2. Il est impossible que deux instructions `return` se trouvent dans le même chemin de code, supprimez donc la dernière ligne `return context.sync();` à la fin de la fonction `Word.run`. Vous ajouterez une nouvelle ligne finale `context.sync` par la suite dans ce didacticiel.

3. Coupez la ligne `doc.body.insertParagraph` et collez-la à la place de `TODO4`.

4. Remplacez `TODO5` par le code suivant. Remarque :

   - Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que la logique `insertParagraph` n’a pas été mise en file d’attente.

   - La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de context.sync.

    ```js
    .then(context.sync);
    ```

Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");
                }
            )
            .then(context.sync);
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

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `insert-text-into-range`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-text-into-range`, ajoutez le code suivant :

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. Sous la fonction `insertTextIntoRange`, ajoutez la fonction suivante :

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {

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

6. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - La méthode est destinée à ajouter une plage dont le texte est « Office 2019 », avant la plage contenant le texte « Office 365 ». Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à ajouter.

   - Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Pour plus d’informations sur les options d’emplacement, reportez-vous à la discussion précédente sur la fonction `insertTextIntoRange`.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. Remplacez `TODO2` par le code suivant.

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.

                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

8. Remplacez `TODO3` par le code suivant. Ce nouveau paragraphe montre que le nouveau texte n’entre ***pas*** dans la plage sélectionnée d’origine. La plage d’origine contient toujours le texte qu’elle contenait lorsqu’elle avait été sélectionnée uniquement.

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. Remplacez `TODO4` par le code suivant :

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a>Remplacer le texte d’une plage

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `insert-text-outside-range`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-text-outside-range`, ajoutez le code suivant :

    ```js
    $('#replace-text').click(replaceText);
    ```

5. Sous la fonction `insertTextBeforeRange`, ajoutez la fonction suivante :

    ```js
    function replaceText() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text.

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

6. Remplacez `TODO1` par le code suivant. La méthode est destinée à remplacer la chaîne « several » (plusieurs) par la chaîne « many » (beaucoup). Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Dans le volet Office, sélectionnez **Insérer un paragraphe** pour vous assurer qu’un paragraphe apparaît au début du document.

6. Sélectionnez du texte. Sélectionner l’expression « Click-to-Run » (Démarrer en un clic) semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*

7. Sélectionnez le bouton **Insérer une abréviation**. L’abréviation « (C2R) » est ajoutée. Notez également qu’en bas du document, un nouveau paragraphe est ajouté avec l’intégralité du texte développé, car la nouvelle chaîne a été ajoutée à la plage existante.

8. Sélectionnez du texte. Sélectionner l’expression « Office 365 » semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*

9. Sélectionnez le bouton **Ajouter les informations de version**. L’expression « Office 2019 » est insérée entre « Office 2016 » et « Office 365 ». Notez également qu’en bas du document, un nouveau paragraphe est ajouté. Celui-ci contient uniquement le texte sélectionné à l’origine, car la nouvelle chaîne est devenue une nouvelle plage plutôt que d’être ajoutée à la plage d’origine.

10. Sélectionnez du texte. Sélectionner le mot « several » (plusieurs) semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*

11. Sélectionnez le bouton permettant de **modifier la condition de quantité** (Change Quantity Term). Notez que « many » (beaucoup) remplace le texte sélectionné.

    ![Didacticiel Word- Ajout et remplacement de texte](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a>Insérer des images, du code HTML et des tableaux

Dans cette étape du didacticiel, vous allez découvrir comment insérer des images, du code HTML et des tableaux dans le document.

### <a name="insert-an-image"></a>Insérer une image

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Dans la partie supérieure du fichier, juste en dessous de la ligne stricte, ajoutez la ligne suivante. Cette ligne importe une variable à partir d’un autre fichier. La variable est une chaîne en base 64 qui encode une image. Pour afficher la chaîne encodée, ouvrez le fichier base64Image.js dans la racine du projet.

    ```js
    import { base64Image } from "./base64Image";
    ```

6. Sous la ligne qui attribue un gestionnaire de clics au bouton `replace-text`, ajoutez le code suivant :

    ```js
    $('#insert-image').click(insertImage);
    ```

7. Sous la fonction `replaceText`, ajoutez la fonction suivante :

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

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

8. Remplacez `TODO1` par le code suivant. Cette ligne insère l’image encodée en base 64 à la fin du document. (L’objet `Paragraph` contient également une méthode `insertInlinePictureFromBase64` et d’autres méthodes `insert*`. Reportez-vous à la section Insérer du code HTML suivante pour consulter un exemple.)

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a>Insérer du code HTML

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `insert-image`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-image`, ajoutez le code suivant :

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. Sous la fonction `insertImage`, ajoutez la fonction suivante :

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

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

6. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - La première ligne ajoute un paragraphe vide à la fin du document. 

   - La deuxième ligne insère une chaîne de code HTML à la fin du paragraphe. Plus précisément, deux paragraphes : un paragraphe avec la police Verdana, et l’autre avec le style par défaut du document Word. (Comme dans la méthode `insertImage` précédente, l’objet `context.document.body` dispose également des méthodes `insert*`.)

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a>Insérer une forme

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `insert-html`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-html`, ajoutez le code suivant :

    ```js
    $('#insert-table').click(insertTable);
    ```

5. Sous la fonction `insertHTML`, ajoutez la fonction suivante :

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

6. Remplacez `TODO1` par le code suivant. Cette ligne utilise la méthode `ParagraphCollection.getFirst` pour obtenir une référence au premier paragraphe, puis utilise la méthode `Paragraph.getNext` pour obtenir une référence au deuxième paragraphe.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. Remplacez `TODO2` par le code suivant. Veuillez noter les informations suivantes :

   - Les deux premiers paramètres de la méthode `insertTable` spécifient le nombre de lignes et de colonnes.

   - Le troisième paramètre indique l’emplacement où insérer le tableau, en l’occurrence après le paragraphe.

   - Le quatrième paramètre est une matrice à deux dimensions qui définit les valeurs des cellules du tableau.

   - Le tableau aura un style par défaut brut, mais la méthode `insertTable` renvoie un objet `Table` avec de nombreux membres, dont certains sont utilisés pour définir le style du tableau.

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Dans le volet Office, sélectionnez **Insérer un paragraphe** au moins trois fois pour vous assurer qu’il existe quelques paragraphes dans le document.

6. Sélectionnez le bouton **Insérer une image** et vous remarquerez qu’une image est insérée à la fin du document.

7. Sélectionnez le bouton **Insérer du code HTML**, puis notez que deux paragraphes sont insérés à la fin du document, et que le premier est affiché dans la police Verdana.

8. Sélectionnez le bouton **Insérer un tableau** et notez qu’un tableau est inséré après le deuxième paragraphe.

    ![Didacticiel Word- Insérer une image, du code HTML et un tableau](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a>Créer et mettre à jour des contrôles de contenu

Dans cette étape du didacticiel, vous découvrirez comment créer des contrôles de contenu de texte enrichi dans le document, puis comment insérer et remplacer du contenu dans les contrôles.

> [!NOTE]
> Il existe plusieurs types de contrôles de contenu pouvant être ajoutés à un document Word via l’interface utilisateur. Toutefois, actuellement, seuls les contrôles de contenu de texte enrichi sont pris en charge par Word.js.
>
> Avant de commencer cette étape du didacticiel, nous vous recommandons de créer et de manipuler des contrôles de contenu de texte enrichi via l’interface utilisateur Word afin de vous familiariser avec les contrôles et leurs propriétés. Pour plus d’informations, reportez-vous à l’article [Créer des formulaires à remplir ou imprimer dans Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).

### <a name="create-a-content-control"></a>Créer un contrôle de contenu

1. Ouvrez le projet dans votre éditeur de code.

2. Ouvrez le fichier index.html.

3. En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. Ouvrez le fichier app.js.

5. Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-table`, ajoutez le code suivant :

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. Sous la fonction `insertTable`, ajoutez la fonction suivante :

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

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

7. Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :

   - Ce code est destiné à intégrer l’expression « Office 365 » dans un contrôle de contenu. Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.

   - La propriété `ContentControl.title` indique le titre visible du contrôle de contenu.

   - La propriété `ContentControl.tag` indique une balise qui peut être utilisée pour obtenir une référence à un contrôle de contenu à l’aide de la méthode `ContentControlCollection.getByTag`, que vous utiliserez dans une fonction ultérieure.

   - La propriété `ContentControl.appearance` indique l’apparence visuelle du contrôle. Utiliser la valeur « Tags » (Balises) signifie que le contrôle est intégré entre des balises de début et de fin, et que la balise de début portera le titre du contrôle de contenu. Les autres valeurs possibles sont « BoundingBox » (Cadre englobant) et « None » (Aucun).

   - La propriété `ContentControl.color` spécifie la couleur des balises ou la bordure du cadre englobant.

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a>Remplacer le contenu du contrôle de contenu

1. Ouvrez le fichier index.html.

2. En dessous de la balise `div` qui contient le bouton `create-content-control`, ajoutez le balisage suivant :

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. Ouvrez le fichier app.js.

4. Sous la ligne qui attribue un gestionnaire de clics au bouton `create-content-control`, ajoutez le code suivant :

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. Sous la fonction `createContentControl`, ajoutez la fonction suivante :

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

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

6. Remplacez `TODO1` par le code suivant. Remarque:

    - La méthode `ContentControlCollection.getByTag` renvoie un élément `ContentControlCollection` comprenant tous les contrôles de contenu de la balise spécifiée. Nous utilisons `getFirst` pour obtenir une référence pour le contrôle souhaité.

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a>Test du complément

1. Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.

     > [!NOTE]
     > Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.

2. Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.

3. Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.

4. Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.

5. Dans le volet des tâches, sélectionnez **Insérer un paragraphe** pour vous assurer qu’il existe un paragraphe contenant « Office 365 » en haut du document.

6. Sélectionnez l’expression « Office 365 » dans le paragraphe que vous venez d’ajouter, puis sélectionnez le bouton **Créer un contrôle de contenu**. L’expression est intégrée dans des balises nommées « Service name » (Nom de service).

7. Sélectionnez le bouton **Renommer le service** et notez que le texte du contrôle de contenu devient « Fabrikam Online Productivity Suite ».

    ![Didacticiel Word-Créer un contrôle de contenu et modifier son texte](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a>Étapes suivantes

Dans ce didacticiel, vous avez créé un Word tâche volet complément qui insère et remplace le texte, images et autres content dans un document Word. Pour en savoir plus sur le développement des complément Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Présentation des compléments Word](../word/word-add-ins-programming-overview.md)
