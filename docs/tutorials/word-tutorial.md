---
title: Didacticiel sur les compléments Word
description: Dans ce didacticiel, vous allez cr?er un compl?ment Word qui ins?re (et remplace) des plages de texte, des paragraphes, des images, du code HTML, des tableaux et des contr?les de contenu. Vous découvrirez également comment mettre en forme du texte et comment insérer (et remplacer) du contenu dans les contrôles de contenu.
ms.date: 06/20/2019
ms.prod: word
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 60397eb4afce60a0880f19be8296ad5fdce315a8
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771869"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a><span data-ttu-id="95026-104">Didacticiel : Créer un complément de volet de tâches Word</span><span class="sxs-lookup"><span data-stu-id="95026-104">Tutorial: Create a Word task pane add-in</span></span>

<span data-ttu-id="95026-105">Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:</span><span class="sxs-lookup"><span data-stu-id="95026-105">In this tutorial, you'll create a Word task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="95026-106">Insère une plage de texte</span><span class="sxs-lookup"><span data-stu-id="95026-106">Inserts a range of text</span></span>
> * <span data-ttu-id="95026-107">Formats de texte</span><span class="sxs-lookup"><span data-stu-id="95026-107">Formats text</span></span>
> * <span data-ttu-id="95026-108">Remplacer du texte et insérer du texte à divers emplacements</span><span class="sxs-lookup"><span data-stu-id="95026-108">Replaces text and inserts text in various locations</span></span>
> * <span data-ttu-id="95026-109">Insère des images, du code HTML et des tableaux</span><span class="sxs-lookup"><span data-stu-id="95026-109">Inserts images, HTML, and tables</span></span>
> * <span data-ttu-id="95026-110">Crée et met à jour des contrôles de contenu</span><span class="sxs-lookup"><span data-stu-id="95026-110">Creates and updates content controls</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="95026-111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="95026-111">Prerequisites</span></span>

<span data-ttu-id="95026-112">Pour utiliser ce didacticiel, les logiciels suivants doivent être installés.</span><span class="sxs-lookup"><span data-stu-id="95026-112">To use this tutorial, you need to have the following installed.</span></span>

- <span data-ttu-id="95026-p102">Word 2016, version 1711 (Démarrer en un clic version 8730.1000) ou version ultérieure. Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="95026-p102">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="95026-116">Node</span><span class="sxs-lookup"><span data-stu-id="95026-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="95026-117">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="95026-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="95026-118">Créer votre projet de complément</span><span class="sxs-lookup"><span data-stu-id="95026-118">Create your add-in project</span></span>

<span data-ttu-id="95026-119">Procédez comme suit pour créer le projet de complément Word que vous souhaitez utiliser comme base pour ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="95026-119">Complete the following steps to create the Word add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="95026-120">Clonez le référentiel GitHub du [didacticiel sur les compléments Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="95026-120">Clone the GitHub repository [Word add-in tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="95026-121">Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="95026-122">Exécutez la commande `npm install` pour installer les outils et les bibliothèques répertoriées dans le fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="95026-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="95026-123">Suivez les étapes de l' [installation du certificat auto-signé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="95026-123">Carry out the steps in [Installing the self-signed certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="insert-a-range-of-text"></a><span data-ttu-id="95026-124">Insérer une plage de texte</span><span class="sxs-lookup"><span data-stu-id="95026-124">Insert a range of text</span></span>

<span data-ttu-id="95026-125">Dans cette étape du tutoriel, vous devez tester par programme que votre complément prend en charge la version actuelle de Word de l’utilisateur, puis insérer un paragraphe dans le document.</span><span class="sxs-lookup"><span data-stu-id="95026-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="95026-126">Codage du complément</span><span class="sxs-lookup"><span data-stu-id="95026-126">Code the add-in</span></span>

1. <span data-ttu-id="95026-127">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="95026-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="95026-128">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-128">Open the file index.html.</span></span>

3. <span data-ttu-id="95026-129">Remplacez `TODO1` par le codage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="95026-130">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-130">Open the app.js file.</span></span>

5. <span data-ttu-id="95026-p103">Remplacez `TODO1` par le code suivant. Ce code détermine si la version de Word de l’utilisateur prend en charge une version de Word.js qui inclut toutes les API utilisées dans les étapes de ce didacticiel. Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge. Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par sa version d’Excel.</span><span class="sxs-lookup"><span data-stu-id="95026-p103">Replace the `TODO1` with the following code. This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="95026-135">Remplacez `TODO2` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="95026-136">Remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="95026-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="95026-137">Remarque :</span><span class="sxs-lookup"><span data-stu-id="95026-137">Note:</span></span>

   - <span data-ttu-id="95026-p105">Votre logique métier Word.js est ajoutée à la fonction qui est transmise à `Word.run`. Cette logique n’est pas exécutée immédiatement. Au lieu de cela, elle est ajoutée à une file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="95026-p105">Your Word.js business logic will be added to the function that is passed to `Word.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="95026-141">La méthode `context.sync` envoie toutes les commandes en file d’attente vers Word pour exécution.</span><span class="sxs-lookup"><span data-stu-id="95026-141">The `context.sync` method sends all queued commands to Word for execution.</span></span>

   - <span data-ttu-id="95026-p106">L’élément `Word.run` est suivi par un bloc `catch`. Il s’agit d’une meilleure pratique que vous devez toujours suivre.</span><span class="sxs-lookup"><span data-stu-id="95026-p106">The `Word.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="95026-p107">Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-p107">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-146">Le premier paramètre de la méthode `insertParagraph` correspond au texte pour le nouveau paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-146">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>

   - <span data-ttu-id="95026-p108">Le deuxième paramètre correspond à l’emplacement dans le corps où sera inséré le paragraphe. Les autres options d’insertion de paragraphe, lorsque l’objet parent est le corps, sont « Fin » et « Remplacer ».</span><span class="sxs-lookup"><span data-stu-id="95026-p108">The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                            "Start");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="95026-149">Test du complément</span><span class="sxs-lookup"><span data-stu-id="95026-149">Test the add-in</span></span>

1. <span data-ttu-id="95026-150">Ouvrez une fenêtre Git Bash ou une invite système activée par Node.JS, et accédez au dossier **Démarrer** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-150">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="95026-151">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="95026-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="95026-152">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="95026-152">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="95026-153">Chargez une version test du complément en utilisant l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-153">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="95026-154">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="95026-154">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="95026-155">Navigateur Web: [chargement de compléments Office dans Office sur le Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="95026-155">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>

    - <span data-ttu-id="95026-156">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="95026-156">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="95026-157">Dans le menu **Accueil** de Word, sélectionnez **Afficher le volet des tâches**.</span><span class="sxs-lookup"><span data-stu-id="95026-157">On the **Home** menu of Word, select **Show Taskpane**.</span></span>

6. <span data-ttu-id="95026-158">Dans le volet Office, sélectionnez **Insérer un paragraphe**.</span><span class="sxs-lookup"><span data-stu-id="95026-158">In the task pane, choose **Insert Paragraph**.</span></span>

7. <span data-ttu-id="95026-159">Apportez une modification au paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-159">Make a change in the paragraph.</span></span>

8. <span data-ttu-id="95026-160">Sélectionnez à nouveau **Insérer un paragraphe**.</span><span class="sxs-lookup"><span data-stu-id="95026-160">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="95026-161">Notez que le nouveau paragraphe se trouve au-dessus du précédent, car la méthode `insertParagraph` effectue l’insertion au « début » du corps du document.</span><span class="sxs-lookup"><span data-stu-id="95026-161">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the start of the document's body.</span></span>

    ![Didacticiel Word- Insérer un paragraphe](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a><span data-ttu-id="95026-163">Mettre en forme du texte</span><span class="sxs-lookup"><span data-stu-id="95026-163">Format text</span></span>

<span data-ttu-id="95026-164">Dans cette étape du didacticiel, vous devez appliquer un style intégré au texte, appliquer un style personnalisé à texte et modifier la police du texte.</span><span class="sxs-lookup"><span data-stu-id="95026-164">In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.</span></span>

### <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="95026-165">Appliquer un style prédéfini au texte</span><span class="sxs-lookup"><span data-stu-id="95026-165">Apply a built-in style to text</span></span>

1. <span data-ttu-id="95026-166">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="95026-166">Open the project in your code editor.</span></span> 

2. <span data-ttu-id="95026-167">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-167">Open the file index.html.</span></span>

3. <span data-ttu-id="95026-168">Juste en dessous de la balise `div` qui contient le bouton `insert-paragraph`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-168">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="95026-169">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-169">Open the app.js file.</span></span>

5. <span data-ttu-id="95026-170">Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-paragraph`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-170">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="95026-171">Ajoutez la fonction suivante juste après la fonction `insertParagraph`:</span><span class="sxs-lookup"><span data-stu-id="95026-171">Just below the `insertParagraph` function, add the following function:</span></span>

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

7. <span data-ttu-id="95026-p110">Remplacez `TODO1` par le code suivant. Le code applique un style à un paragraphe, mais les styles peuvent également être appliqués aux plages de texte.</span><span class="sxs-lookup"><span data-stu-id="95026-p110">Replace `TODO1` with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="95026-174">Appliquer un style personnalisé au texte</span><span class="sxs-lookup"><span data-stu-id="95026-174">Apply a custom style to text</span></span>

1. <span data-ttu-id="95026-175">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-175">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-176">En dessous de la balise `div` qui contient le bouton `apply-style`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-176">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="95026-177">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-177">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-178">Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-style`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-178">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="95026-179">Sous la fonction `applyStyle`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-179">Below the `applyStyle` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p111">Remplacez `TODO1` par le code suivant. Le code applique un style personnalisé qui n’existe pas encore. Vous allez créer un style nommé **MyCustomStyle** lors de l’étape [Test du complément](#test-the-add-in).</span><span class="sxs-lookup"><span data-stu-id="95026-p111">Replace `TODO1` with the following code. Note that the code applies a custom style that does not exist yet. You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a><span data-ttu-id="95026-183">Modifier la police du texte</span><span class="sxs-lookup"><span data-stu-id="95026-183">Change the font of text</span></span>

1. <span data-ttu-id="95026-184">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-184">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-185">En dessous de la balise `div` qui contient le bouton `apply-custom-style`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-185">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="95026-186">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-186">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-187">Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-custom-style`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-187">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="95026-188">Sous la fonction `applyCustomStyle`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-188">Below the `applyCustomStyle` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p112">Remplacez `TODO1` par le code suivant. Le code obtient une référence au deuxième paragraphe en utilisant la méthode `ParagraphCollection.getFirst` chaînée à la méthode `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="95026-p112">Replace `TODO1` with the following code. Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a><span data-ttu-id="95026-191">Test du complément</span><span class="sxs-lookup"><span data-stu-id="95026-191">Test the add-in</span></span>

1. <span data-ttu-id="95026-192">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="95026-192">In the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="95026-193">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-193">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="95026-p114">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="95026-p114">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="95026-198">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="95026-198">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="95026-199">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="95026-199">Run the command `npm start` to start a web server running on localhost.</span></span>   

4. <span data-ttu-id="95026-200">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="95026-200">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="95026-p115">Assurez-vous qu’il existe au moins trois paragraphes dans le document. Vous pouvez sélectionner trois fois l’option **Insérer un paragraphe**. *Vérifiez attentivement qu’aucun paragraphe vide n’apparaît à la fin du document. S’il y en a un, supprimez-le.*</span><span class="sxs-lookup"><span data-stu-id="95026-p115">Be sure there are at least three paragraphs in the document. You can choose **Insert Paragraph** three times. *Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>

6. <span data-ttu-id="95026-p116">Dans Word, créez un style personnalisé nommé « MyCustomStyle ». Vous pouvez y appliquer la mise en forme que vous souhaitez.</span><span class="sxs-lookup"><span data-stu-id="95026-p116">In Word, create a custom style named "MyCustomStyle". It can have any formatting that you want.</span></span>

7. <span data-ttu-id="95026-p117">Sélectionnez le bouton **Appliquer le style**. Le style prédéfini **Référence intense** est appliqué au premier paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-p117">Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>

8. <span data-ttu-id="95026-p118">Sélectionnez le bouton **Appliquer un style personnalisé**. Votre style personnalisé est appliqué au dernier paragraphe. (Si rien ne semble se produire, le dernier paragraphe est peut-être vide. Si c’est le cas, ajoutez-y du texte.)</span><span class="sxs-lookup"><span data-stu-id="95026-p118">Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)</span></span>

9. <span data-ttu-id="95026-p119">Sélectionnez le bouton **Modifier la police**. La police Courier New, 18 pt, en gras, est appliquée au deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-p119">Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Didacticiel Word- Appliquer des styles et une police](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a><span data-ttu-id="95026-215">Remplacer du texte et insérer du texte</span><span class="sxs-lookup"><span data-stu-id="95026-215">Replace text and insert text</span></span>

<span data-ttu-id="95026-216">Dans cette étape du didacticiel, vous ajouterez du texte dans les plages de texte sélectionnées et en dehors de celles-ci, puis remplacerez le texte de la plage sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="95026-216">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span>

### <a name="add-text-inside-a-range"></a><span data-ttu-id="95026-217">Ajouter du texte dans une plage</span><span class="sxs-lookup"><span data-stu-id="95026-217">Add text inside a range</span></span>

1. <span data-ttu-id="95026-218">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="95026-218">Open the project in your code editor.</span></span>

2. <span data-ttu-id="95026-219">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-219">Open the file index.html.</span></span>

3. <span data-ttu-id="95026-220">En dessous de la balise `div` qui contient le bouton `change-font`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-220">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. <span data-ttu-id="95026-221">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-221">Open the app.js file.</span></span>

5. <span data-ttu-id="95026-222">Sous la ligne qui attribue un gestionnaire de clics au bouton `change-font`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-222">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="95026-223">Sous la fonction `changeFont`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-223">Below the `changeFont` function, add the following function:</span></span>

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

7. <span data-ttu-id="95026-p120">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-p120">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-p121">La méthode est destinée à insérer l’abréviation [« (C2R) »] à la fin de la plage dont le texte est « Click-to-Run » (Démarrer en un clic). Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="95026-p121">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="95026-228">Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à insérer dans l’objet `Range`.</span><span class="sxs-lookup"><span data-stu-id="95026-228">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>

   - <span data-ttu-id="95026-p122">Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Outre « Fin », les autres options possibles sont : « Début », « Avant », « Après » et « Remplacer ».</span><span class="sxs-lookup"><span data-stu-id="95026-p122">The second parameter specifies where in the range the additional text should be inserted. Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 

   - <span data-ttu-id="95026-p123">La différence entre « Fin » et « Après » est que « Fin » insère le nouveau texte à la fin de la plage existante, tandis que l’option « Après » crée une plage avec la chaîne et insère la nouvelle plage après la plage existante. De même, « Début » insère le texte au début de la plage existante, tandis que l’option « Avant » insère une nouvelle plage. L’option « Remplacer » remplace le texte de la plage existante par la chaîne dans le premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="95026-p123">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range. Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range. "Replace" replaces the text of the existing range with the string in the first parameter.</span></span>

   - <span data-ttu-id="95026-p124">Vous avez vu lors d’une étape précédente du didacticiel que les méthodes insert\* de l’objet corps ne disposent pas des options « Avant » et « Après ». Cela est dû au fait que vous ne pouvez pas placer de contenu en dehors du corps du document.</span><span class="sxs-lookup"><span data-stu-id="95026-p124">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options. This is because you can't put content outside of the document's body.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. <span data-ttu-id="95026-p125">Nous ignorerons `TODO2` jusqu’à la section suivante. Remplacez `TODO3` par le code suivant. Ce code est similaire au code que vous avez créé lors de la première phase du didacticiel, sauf que, maintenant, vous insérez un nouveau paragraphe à la fin du document plutôt qu’au début. Ce nouveau paragraphe montre que le nouveau texte fait désormais partie de la plage d’origine.</span><span class="sxs-lookup"><span data-stu-id="95026-p125">We'll skip over `TODO2` until the next section. Replace `TODO3` with the following code. This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start. This new paragraph will demonstrate that the new text is now part of the original range.</span></span>

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="95026-240">Ajouter du code pour récupérer des propriétés de document dans les objets de script du volet Office</span><span class="sxs-lookup"><span data-stu-id="95026-240">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="95026-p126">Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour écrire (*write*) dans le document Office. Chaque fonction se terminait par un appel de la méthode `context.sync()` qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées. Cependant, le code que vous avez ajouté dans la dernière étape appelle la propriété `originalRange.text` et c’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `originalRange` est uniquement un objet de proxy qui existe dans le script de votre volet Office. Il ne connaît pas le texte réel de la plage dans le document, donc sa propriété `text` ne peut pas contenir de valeur réelle. Il est nécessaire de récupérer d’abord la valeur de texte de la plage à partir du document, puis de l’utiliser pour définir la valeur de `originalRange.text`. Seulement ensuite, la propriété `originalRange.text` peut être appelée sans générer d’exception. Ce processus de récupération comporte trois étapes :</span><span class="sxs-lookup"><span data-stu-id="95026-p126">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value. It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`. Only then can `originalRange.text` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="95026-248">Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.</span><span class="sxs-lookup"><span data-stu-id="95026-248">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="95026-249">Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.</span><span class="sxs-lookup"><span data-stu-id="95026-249">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="95026-250">Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.</span><span class="sxs-lookup"><span data-stu-id="95026-250">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="95026-251">Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.</span><span class="sxs-lookup"><span data-stu-id="95026-251">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="95026-252">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="95026-252">Replace `TODO2` with the following code.</span></span>
  
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

2. <span data-ttu-id="95026-p127">Il est impossible que deux instructions `return` se trouvent dans le même chemin de code, supprimez donc la dernière ligne `return context.sync();` à la fin de la fonction `Word.run`. Vous ajouterez une nouvelle ligne finale `context.sync` par la suite dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="95026-p127">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span>

3. <span data-ttu-id="95026-255">Coupez la ligne `doc.body.insertParagraph` et collez-la à la place de `TODO4`.</span><span class="sxs-lookup"><span data-stu-id="95026-255">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span>

4. <span data-ttu-id="95026-p128">Remplacez `TODO5` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="95026-p128">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-258">Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que la logique `insertParagraph` n’a pas été mise en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="95026-258">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>

   - <span data-ttu-id="95026-259">La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de context.sync.</span><span class="sxs-lookup"><span data-stu-id="95026-259">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="95026-260">Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="95026-260">When you're done, the entire function should look like the following:</span></span>

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

### <a name="add-text-between-ranges"></a><span data-ttu-id="95026-261">Ajouter du texte entre les plages</span><span class="sxs-lookup"><span data-stu-id="95026-261">Add text between ranges</span></span>

1. <span data-ttu-id="95026-262">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-262">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-263">En dessous de la balise `div` qui contient le bouton `insert-text-into-range`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-263">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. <span data-ttu-id="95026-264">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-264">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-265">Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-text-into-range`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-265">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="95026-266">Sous la fonction `insertTextIntoRange`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-266">Below the `insertTextIntoRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p129">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-p129">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-p130">La méthode est destinée à ajouter une plage dont le texte est « Office 2019 », avant la plage contenant le texte « Office 365 ». Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="95026-p130">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="95026-271">Le premier paramètre de la méthode `Range.insertText` correspond à la chaîne à ajouter.</span><span class="sxs-lookup"><span data-stu-id="95026-271">The first parameter of the `Range.insertText` method is the string to add.</span></span>

   - <span data-ttu-id="95026-p131">Le deuxième paramètre spécifie l’emplacement où le texte supplémentaire doit être inséré dans la plage. Pour plus d’informations sur les options d’emplacement, reportez-vous à la discussion précédente sur la fonction `insertTextIntoRange`.</span><span class="sxs-lookup"><span data-stu-id="95026-p131">The second parameter specifies where in the range the additional text should be inserted. For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. <span data-ttu-id="95026-274">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="95026-274">Replace `TODO2` with the following code.</span></span>

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

8. <span data-ttu-id="95026-p132">Remplacez `TODO3` par le code suivant. Ce nouveau paragraphe montre que le nouveau texte n’entre ***pas*** dans la plage sélectionnée d’origine. La plage d’origine contient toujours le texte qu’elle contenait lorsqu’elle avait été sélectionnée uniquement.</span><span class="sxs-lookup"><span data-stu-id="95026-p132">Replace `TODO3` with the following code. This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range. The original range still has only the text it had when it was selected.</span></span>

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. <span data-ttu-id="95026-278">Remplacez `TODO4` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-278">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a><span data-ttu-id="95026-279">Remplacer le texte d’une plage</span><span class="sxs-lookup"><span data-stu-id="95026-279">Replace the text of a range</span></span>

1. <span data-ttu-id="95026-280">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-280">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-281">En dessous de la balise `div` qui contient le bouton `insert-text-outside-range`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-281">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. <span data-ttu-id="95026-282">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-282">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-283">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-text-outside-range`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-283">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="95026-284">Sous la fonction `insertTextBeforeRange`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-284">Below the `insertTextBeforeRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p133">Remplacez `TODO1` par le code suivant. La méthode est destinée à remplacer la chaîne « several » (plusieurs) par la chaîne « many » (beaucoup). Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="95026-p133">Replace `TODO1` with the following code. Note that the method is intended to replace the string "several" with the string "many". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="95026-288">Test du complément</span><span class="sxs-lookup"><span data-stu-id="95026-288">Test the add-in</span></span>

1. <span data-ttu-id="95026-p134">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-p134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="95026-p135">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="95026-p135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="95026-295">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="95026-295">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="95026-296">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="95026-296">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="95026-297">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="95026-297">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="95026-298">Dans le volet Office, sélectionnez **Insérer un paragraphe** pour vous assurer qu’un paragraphe apparaît au début du document.</span><span class="sxs-lookup"><span data-stu-id="95026-298">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>

6. <span data-ttu-id="95026-p136">Sélectionnez du texte. Sélectionner l’expression « Click-to-Run » (Démarrer en un clic) semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*</span><span class="sxs-lookup"><span data-stu-id="95026-p136">Select some text. Selecting the phrase "Click-to-Run" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

7. <span data-ttu-id="95026-p137">Sélectionnez le bouton **Insérer une abréviation**. L’abréviation « (C2R) » est ajoutée. Notez également qu’en bas du document, un nouveau paragraphe est ajouté avec l’intégralité du texte développé, car la nouvelle chaîne a été ajoutée à la plage existante.</span><span class="sxs-lookup"><span data-stu-id="95026-p137">Choose the **Insert Abbreviation** button. Note that " (C2R)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>

8. <span data-ttu-id="95026-p138">Sélectionnez du texte. Sélectionner l’expression « Office 365 » semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*</span><span class="sxs-lookup"><span data-stu-id="95026-p138">Select some text. Selecting the phrase "Office 365" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

9. <span data-ttu-id="95026-p139">Sélectionnez le bouton **Ajouter les informations de version**. L’expression « Office 2019 » est insérée entre « Office 2016 » et « Office 365 ». Notez également qu’en bas du document, un nouveau paragraphe est ajouté. Celui-ci contient uniquement le texte sélectionné à l’origine, car la nouvelle chaîne est devenue une nouvelle plage plutôt que d’être ajoutée à la plage d’origine.</span><span class="sxs-lookup"><span data-stu-id="95026-p139">Choose the **Add Version Info** button. Note that "Office 2019, " is inserted between "Office 2016" and "Office 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>

10. <span data-ttu-id="95026-p140">Sélectionnez du texte. Sélectionner le mot « several » (plusieurs) semble le plus approprié. *Veillez à ne pas inclure tout espace précédent ou suivant dans la sélection.*</span><span class="sxs-lookup"><span data-stu-id="95026-p140">Select some text. Selecting the word "several" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

11. <span data-ttu-id="95026-p141">Sélectionnez le bouton permettant de **modifier la condition de quantité** (Change Quantity Term). Notez que « many » (beaucoup) remplace le texte sélectionné.</span><span class="sxs-lookup"><span data-stu-id="95026-p141">Choose the **Change Quantity Term** button. Note that "many" replaces the selected text.</span></span>

    ![Didacticiel Word- Ajout et remplacement de texte](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a><span data-ttu-id="95026-317">Insérer des images, du code HTML et des tableaux</span><span class="sxs-lookup"><span data-stu-id="95026-317">Insert images, HTML, and tables</span></span>

<span data-ttu-id="95026-318">Dans cette étape du didacticiel, vous allez découvrir comment insérer des images, du code HTML et des tableaux dans le document.</span><span class="sxs-lookup"><span data-stu-id="95026-318">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

### <a name="insert-an-image"></a><span data-ttu-id="95026-319">Insérer une image</span><span class="sxs-lookup"><span data-stu-id="95026-319">Insert an image</span></span>

1. <span data-ttu-id="95026-320">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="95026-320">Open the project in your code editor.</span></span>

2. <span data-ttu-id="95026-321">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-321">Open the file index.html.</span></span>

3. <span data-ttu-id="95026-322">En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-322">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="95026-323">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-323">Open the app.js file.</span></span>

5. <span data-ttu-id="95026-p142">Dans la partie supérieure du fichier, juste en dessous de la ligne stricte, ajoutez la ligne suivante. Cette ligne importe une variable à partir d’un autre fichier. La variable est une chaîne en base 64 qui encode une image. Pour afficher la chaîne encodée, ouvrez le fichier base64Image.js dans la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-p142">Near the top of the file, just below the use-strict line, add the following line. This line imports a variable from another file. The variable is a base 64 string that encodes an image. To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="95026-328">Sous la ligne qui attribue un gestionnaire de clics au bouton `replace-text`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-328">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="95026-329">Sous la fonction `replaceText`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-329">Below the `replaceText` function, add the following function:</span></span>

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

8. <span data-ttu-id="95026-p143">Remplacez `TODO1` par le code suivant. Cette ligne insère l’image encodée en base 64 à la fin du document. (L’objet `Paragraph` contient également une méthode `insertInlinePictureFromBase64` et d’autres méthodes `insert*`. Reportez-vous à la section Insérer du code HTML suivante pour consulter un exemple.)</span><span class="sxs-lookup"><span data-stu-id="95026-p143">Replace `TODO1` with the following code. Note that this line inserts the base 64 encoded image at the end of the document. (The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods. See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a><span data-ttu-id="95026-334">Insérer du code HTML</span><span class="sxs-lookup"><span data-stu-id="95026-334">Insert HTML</span></span>

1. <span data-ttu-id="95026-335">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-335">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-336">En dessous de la balise `div` qui contient le bouton `insert-image`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-336">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="95026-337">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-337">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-338">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-image`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-338">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="95026-339">Sous la fonction `insertImage`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-339">Below the `insertImage` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p144">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-p144">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-342">La première ligne ajoute un paragraphe vide à la fin du document.</span><span class="sxs-lookup"><span data-stu-id="95026-342">The first line adds a blank paragraph to the end of the document.</span></span> 

   - <span data-ttu-id="95026-p145">La deuxième ligne insère une chaîne de code HTML à la fin du paragraphe. Plus précisément, deux paragraphes : un paragraphe avec la police Verdana, et l’autre avec le style par défaut du document Word. (Comme pour la méthode `insertImage` précédente, l’objet `context.document.body` contient également les méthodes `insert*`.)</span><span class="sxs-lookup"><span data-stu-id="95026-p145">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document. (As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a><span data-ttu-id="95026-345">Insérer une forme</span><span class="sxs-lookup"><span data-stu-id="95026-345">Insert a table</span></span>

1. <span data-ttu-id="95026-346">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-346">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-347">En dessous de la balise `div` qui contient le bouton `insert-html`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-347">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="95026-348">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-348">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-349">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-html`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-349">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="95026-350">Sous la fonction `insertHTML`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-350">Below the `insertHTML` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p146">Remplacez `TODO1` par le code suivant. Cette ligne utilise la méthode `ParagraphCollection.getFirst` pour obtenir une référence au premier paragraphe, puis utilise la méthode `Paragraph.getNext` pour obtenir une référence au deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-p146">Replace `TODO1` with the following code. Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="95026-353">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="95026-353">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="95026-354">Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-354">Note:</span></span>

   - <span data-ttu-id="95026-355">Les deux premiers paramètres de la méthode `insertTable` spécifient le nombre de lignes et de colonnes.</span><span class="sxs-lookup"><span data-stu-id="95026-355">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>

   - <span data-ttu-id="95026-356">Le troisième paramètre indique l’emplacement où insérer le tableau, en l’occurrence après le paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-356">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>

   - <span data-ttu-id="95026-357">Le quatrième paramètre est une matrice à deux dimensions qui définit les valeurs des cellules du tableau.</span><span class="sxs-lookup"><span data-stu-id="95026-357">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>

   - <span data-ttu-id="95026-358">Le tableau aura un style par défaut brut, mais la méthode `insertTable` renvoie un objet `Table` avec de nombreux membres, dont certains sont utilisés pour définir le style du tableau.</span><span class="sxs-lookup"><span data-stu-id="95026-358">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="95026-359">Test du complément</span><span class="sxs-lookup"><span data-stu-id="95026-359">Test the add-in</span></span>

1. <span data-ttu-id="95026-p148">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-p148">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="95026-p149">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="95026-p149">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="95026-366">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="95026-366">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="95026-367">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="95026-367">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="95026-368">Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="95026-368">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="95026-369">Dans le volet Office, sélectionnez **Insérer un paragraphe** au moins trois fois pour vous assurer qu’il existe quelques paragraphes dans le document.</span><span class="sxs-lookup"><span data-stu-id="95026-369">In the task pane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>

6. <span data-ttu-id="95026-370">Sélectionnez le bouton **Insérer une image** et vous remarquerez qu’une image est insérée à la fin du document.</span><span class="sxs-lookup"><span data-stu-id="95026-370">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>

7. <span data-ttu-id="95026-371">Sélectionnez le bouton **Insérer du code HTML**, puis notez que deux paragraphes sont insérés à la fin du document, et que le premier est affiché dans la police Verdana.</span><span class="sxs-lookup"><span data-stu-id="95026-371">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>

8. <span data-ttu-id="95026-372">Sélectionnez le bouton **Insérer un tableau** et notez qu’un tableau est inséré après le deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="95026-372">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Didacticiel Word- Insérer une image, du code HTML et un tableau](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a><span data-ttu-id="95026-374">Créer et mettre à jour des contrôles de contenu</span><span class="sxs-lookup"><span data-stu-id="95026-374">Create and update content controls</span></span>

<span data-ttu-id="95026-375">Dans cette étape du didacticiel, vous découvrirez comment créer des contrôles de contenu de texte enrichi dans le document, puis comment insérer et remplacer du contenu dans les contrôles.</span><span class="sxs-lookup"><span data-stu-id="95026-375">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span>

> [!NOTE]
> <span data-ttu-id="95026-376">Il existe plusieurs types de contrôles de contenu pouvant être ajoutés à un document Word via l’interface utilisateur. Toutefois, actuellement, seuls les contrôles de contenu de texte enrichi sont pris en charge par Word.js.</span><span class="sxs-lookup"><span data-stu-id="95026-376">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>
>
> <span data-ttu-id="95026-p150">Avant de commencer cette étape du didacticiel, nous vous recommandons de créer et de manipuler des contrôles de contenu de texte enrichi via l’interface utilisateur Word afin de vous familiariser avec les contrôles et leurs propriétés. Pour plus d’informations, reportez-vous à l’article [Créer des formulaires à remplir ou imprimer dans Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="95026-p150">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties. For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

### <a name="create-a-content-control"></a><span data-ttu-id="95026-379">Créer un contrôle de contenu</span><span class="sxs-lookup"><span data-stu-id="95026-379">Create a content control</span></span>

1. <span data-ttu-id="95026-380">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="95026-380">Open the project in your code editor.</span></span>

2. <span data-ttu-id="95026-381">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-381">Open the file index.html.</span></span>

3. <span data-ttu-id="95026-382">En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-382">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. <span data-ttu-id="95026-383">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-383">Open the app.js file.</span></span>

5. <span data-ttu-id="95026-384">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-384">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="95026-385">Sous la fonction `insertTable`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-385">Below the `insertTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="95026-p151">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="95026-p151">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="95026-p152">Ce code est destiné à intégrer l’expression « Office 365 » dans un contrôle de contenu. Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="95026-p152">This code is intended to wrap the phrase "Office 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="95026-390">La propriété `ContentControl.title` indique le titre visible du contrôle de contenu.</span><span class="sxs-lookup"><span data-stu-id="95026-390">The `ContentControl.title` property specifies the visible title of the content control.</span></span>

   - <span data-ttu-id="95026-391">La propriété `ContentControl.tag` indique une balise qui peut être utilisée pour obtenir une référence à un contrôle de contenu à l’aide de la méthode `ContentControlCollection.getByTag`, que vous utiliserez dans une fonction ultérieure.</span><span class="sxs-lookup"><span data-stu-id="95026-391">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span>

   - <span data-ttu-id="95026-p153">La propriété `ContentControl.appearance` indique l’apparence visuelle du contrôle. Utiliser la valeur « Tags » (Balises) signifie que le contrôle est intégré entre des balises de début et de fin, et que la balise de début portera le titre du contrôle de contenu. Les autres valeurs possibles sont « BoundingBox » (Cadre englobant) et « None » (Aucun).</span><span class="sxs-lookup"><span data-stu-id="95026-p153">The `ContentControl.appearance` property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".</span></span>

   - <span data-ttu-id="95026-395">La propriété `ContentControl.color` spécifie la couleur des balises ou la bordure du cadre englobant.</span><span class="sxs-lookup"><span data-stu-id="95026-395">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="95026-396">Remplacer le contenu du contrôle de contenu</span><span class="sxs-lookup"><span data-stu-id="95026-396">Replace the content of the content control</span></span>

1. <span data-ttu-id="95026-397">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="95026-397">Open the file index.html.</span></span>

2. <span data-ttu-id="95026-398">En dessous de la balise `div` qui contient le bouton `create-content-control`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-398">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. <span data-ttu-id="95026-399">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="95026-399">Open the app.js file.</span></span>

4. <span data-ttu-id="95026-400">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-content-control`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-400">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="95026-401">Sous la fonction `createContentControl`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="95026-401">Below the `createContentControl` function, add the following function:</span></span>

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

6. <span data-ttu-id="95026-p154">Remplacez `TODO1` par le code suivant. Remarque:</span><span class="sxs-lookup"><span data-stu-id="95026-p154">Replace `TODO1` with the following code. Note:</span></span>

    - <span data-ttu-id="95026-404">La méthode `ContentControlCollection.getByTag` renvoie un élément `ContentControlCollection` comprenant tous les contrôles de contenu de la balise spécifiée.</span><span class="sxs-lookup"><span data-stu-id="95026-404">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="95026-405">Nous utilisons `getFirst` pour obtenir une référence pour le contrôle souhaité.</span><span class="sxs-lookup"><span data-stu-id="95026-405">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="95026-406">Test du complément</span><span class="sxs-lookup"><span data-stu-id="95026-406">Test the add-in</span></span>

1. <span data-ttu-id="95026-p156">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution. Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="95026-p156">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="95026-p157">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet. Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build. Après la commande build, redémarrez le serveur. Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="95026-p157">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="95026-413">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="95026-413">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="95026-414">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="95026-414">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="95026-415">Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="95026-415">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="95026-416">Dans le volet des tâches, sélectionnez **Insérer un paragraphe** pour vous assurer qu’il existe un paragraphe contenant « Office 365 » en haut du document.</span><span class="sxs-lookup"><span data-stu-id="95026-416">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>

6. <span data-ttu-id="95026-p158">Sélectionnez l’expression « Office 365 » dans le paragraphe que vous venez d’ajouter, puis sélectionnez le bouton **Créer un contrôle de contenu**. L’expression est intégrée dans des balises nommées « Service name » (Nom de service).</span><span class="sxs-lookup"><span data-stu-id="95026-p158">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button. Note that the phrase is wrapped in tags labelled "Service Name".</span></span>

7. <span data-ttu-id="95026-419">Sélectionnez le bouton **Renommer le service** et notez que le texte du contrôle de contenu devient « Fabrikam Online Productivity Suite ».</span><span class="sxs-lookup"><span data-stu-id="95026-419">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Didacticiel Word-Créer un contrôle de contenu et modifier son texte](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a><span data-ttu-id="95026-421">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="95026-421">Next steps</span></span>

<span data-ttu-id="95026-422">Dans ce didacticiel, vous avez créé un Word tâche volet complément qui insère et remplace le texte, images et autres content dans un document Word.</span><span class="sxs-lookup"><span data-stu-id="95026-422">In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document.</span></span> <span data-ttu-id="95026-423">Pour en savoir plus sur le développement des complément Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="95026-423">To learn more about building Word add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="95026-424">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="95026-424">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
