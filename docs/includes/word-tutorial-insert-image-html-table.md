<span data-ttu-id="fa75e-101">Dans cette étape du didacticiel, vous allez découvrir comment insérer des images, du code HTML et des tableaux dans le document.</span><span class="sxs-lookup"><span data-stu-id="fa75e-101">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

> [!NOTE]
> <span data-ttu-id="fa75e-p101">Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="fa75e-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="insert-an-image"></a><span data-ttu-id="fa75e-104">Insérer une image</span><span class="sxs-lookup"><span data-stu-id="fa75e-104">Insert an image</span></span>

1. <span data-ttu-id="fa75e-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="fa75e-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="fa75e-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="fa75e-106">Open the file index.html.</span></span>
3. <span data-ttu-id="fa75e-107">En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-107">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="fa75e-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="fa75e-108">Open the app.js file.</span></span>

5. <span data-ttu-id="fa75e-109">Dans la partie supérieure du fichier, juste en dessous de la ligne stricte, ajoutez la ligne suivante.</span><span class="sxs-lookup"><span data-stu-id="fa75e-109">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="fa75e-110">Cette ligne importe une variable à partir d’un autre fichier.</span><span class="sxs-lookup"><span data-stu-id="fa75e-110">This line imports a variable from another file.</span></span> <span data-ttu-id="fa75e-111">La variable est une chaîne en base 64 qui encode une image.</span><span class="sxs-lookup"><span data-stu-id="fa75e-111">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="fa75e-112">Pour afficher la chaîne encodée, ouvrez le fichier base64Image.js dans la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="fa75e-112">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="fa75e-113">Sous la ligne qui attribue un gestionnaire de clics au bouton `replace-text`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-113">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="fa75e-114">Sous la fonction `replaceText`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="fa75e-114">Below the `replaceText` function, add the following function:</span></span>

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

8. <span data-ttu-id="fa75e-115">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="fa75e-115">Replace `TODO1` with the following code:</span></span> <span data-ttu-id="fa75e-116">Cette ligne insère l’image encodée en base 64 à la fin du document.</span><span class="sxs-lookup"><span data-stu-id="fa75e-116">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="fa75e-117">(L’objet `Paragraph` contient également une méthode `insertInlinePictureFromBase64` et d’autres méthodes `insert*`.</span><span class="sxs-lookup"><span data-stu-id="fa75e-117">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="fa75e-118">Reportez-vous à la section Insérer du code HTML suivante pour consulter un exemple.)</span><span class="sxs-lookup"><span data-stu-id="fa75e-118">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

## <a name="insert-html"></a><span data-ttu-id="fa75e-119">Insérer du code HTML</span><span class="sxs-lookup"><span data-stu-id="fa75e-119">Insert HTML</span></span>

1. <span data-ttu-id="fa75e-120">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="fa75e-120">Open the file index.html.</span></span>
2. <span data-ttu-id="fa75e-121">En dessous de la balise `div` qui contient le bouton `insert-image`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-121">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="fa75e-122">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="fa75e-122">Open the app.js file.</span></span>

4. <span data-ttu-id="fa75e-123">Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-image`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-123">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="fa75e-124">Sous la fonction `insertImage`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="fa75e-124">Below the `insertImage` function, add the following function:</span></span>

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

6. <span data-ttu-id="fa75e-p104">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="fa75e-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="fa75e-127">La première ligne ajoute un paragraphe vide à la fin du document.</span><span class="sxs-lookup"><span data-stu-id="fa75e-127">The first line adds a blank paragraph to the end of the document.</span></span> 
   - <span data-ttu-id="fa75e-128">La deuxième ligne insère une chaîne de code HTML à la fin du paragraphe. Plus précisément, deux paragraphes : un paragraphe avec la police Verdana, et l’autre avec le style par défaut du document Word.</span><span class="sxs-lookup"><span data-stu-id="fa75e-128">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="fa75e-129">(Comme pour la méthode `insertImage` précédente, l’objet `context.document.body` contient également les méthodes `insert*`.)</span><span class="sxs-lookup"><span data-stu-id="fa75e-129">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

## <a name="insert-table"></a><span data-ttu-id="fa75e-130">Insérer un tableau</span><span class="sxs-lookup"><span data-stu-id="fa75e-130">Insert Table</span></span>

1. <span data-ttu-id="fa75e-131">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="fa75e-131">Open the file index.html.</span></span>
2. <span data-ttu-id="fa75e-132">En dessous de la balise `div` qui contient le bouton `insert-html`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-132">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="fa75e-133">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="fa75e-133">Open the app.js file.</span></span>

4. <span data-ttu-id="fa75e-134">Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-html`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="fa75e-134">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="fa75e-135">Sous la fonction `insertHTML`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="fa75e-135">Below the `insertHTML` function, add the following function:</span></span>

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

6. <span data-ttu-id="fa75e-136">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="fa75e-136">Replace `TODO1` with the following code:</span></span> <span data-ttu-id="fa75e-137">Cette ligne utilise la méthode `ParagraphCollection.getFirst` pour obtenir une référence au premier paragraphe, puis utilise la méthode `Paragraph.getNext` pour obtenir une référence au deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="fa75e-137">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="fa75e-p107">Remplacez `TODO2` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="fa75e-p107">Replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="fa75e-140">Les deux premiers paramètres de la méthode `insertTable` spécifient le nombre de lignes et de colonnes.</span><span class="sxs-lookup"><span data-stu-id="fa75e-140">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>
   - <span data-ttu-id="fa75e-141">Le troisième paramètre indique l’emplacement où insérer le tableau, en l’occurrence après le paragraphe.</span><span class="sxs-lookup"><span data-stu-id="fa75e-141">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>
   - <span data-ttu-id="fa75e-142">Le quatrième paramètre est une matrice à deux dimensions qui définit les valeurs des cellules du tableau.</span><span class="sxs-lookup"><span data-stu-id="fa75e-142">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>
   - <span data-ttu-id="fa75e-143">Le tableau aura un style par défaut brut, mais la méthode `insertTable` renvoie un objet `Table` avec de nombreux membres, dont certains sont utilisés pour définir le style du tableau.</span><span class="sxs-lookup"><span data-stu-id="fa75e-143">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="fa75e-144">Test du complément</span><span class="sxs-lookup"><span data-stu-id="fa75e-144">Test the add-in</span></span>


1. <span data-ttu-id="fa75e-145">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="fa75e-145">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="fa75e-146">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="fa75e-146">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="fa75e-147">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="fa75e-147">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="fa75e-148">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="fa75e-148">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="fa75e-149">Après la commande build, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="fa75e-149">After the build, restart the server.</span></span> <span data-ttu-id="fa75e-150">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="fa75e-150">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="fa75e-151">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="fa75e-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="fa75e-152">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="fa75e-152">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="fa75e-153">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="fa75e-153">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="fa75e-154">Dans le volet Office, sélectionnez **Insérer un paragraphe** au moins trois fois pour vous assurer qu’il existe quelques paragraphes dans le document.</span><span class="sxs-lookup"><span data-stu-id="fa75e-154">In the taskpane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>
6. <span data-ttu-id="fa75e-155">Sélectionnez le bouton **Insérer une image** et vous remarquerez qu’une image est insérée à la fin du document.</span><span class="sxs-lookup"><span data-stu-id="fa75e-155">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>
7. <span data-ttu-id="fa75e-156">Sélectionnez le bouton **Insérer du code HTML**, puis notez que deux paragraphes sont insérés à la fin du document, et que le premier est affiché dans la police Verdana.</span><span class="sxs-lookup"><span data-stu-id="fa75e-156">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>
8. <span data-ttu-id="fa75e-157">Sélectionnez le bouton **Insérer un tableau** et notez qu’un tableau est inséré après le deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="fa75e-157">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Didacticiel Word - Insérer une image, du code HTML et un tableau](../images/word-tutorial-insert-image-html-table.png)
