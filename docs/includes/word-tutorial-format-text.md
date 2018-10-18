<span data-ttu-id="556c0-101">Dans cette étape du didacticiel, vous modifierez la police du texte, et utiliserez des styles prédéfinis et personnalisés pour le texte.</span><span class="sxs-lookup"><span data-stu-id="556c0-101">In this step of the tutorial, you'll change the font of text, and use both built-in and custom styles on the text.</span></span>

> [!NOTE]
> <span data-ttu-id="556c0-p101">Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="556c0-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="556c0-104">Appliquer un style prédéfini au texte</span><span class="sxs-lookup"><span data-stu-id="556c0-104">Apply a built-in style to text</span></span>

1. <span data-ttu-id="556c0-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="556c0-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="556c0-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="556c0-106">Open the file index.html.</span></span>
3. <span data-ttu-id="556c0-107">Juste en dessous de la balise `div` qui contient le bouton `insert-paragraph`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-107">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="556c0-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="556c0-108">Open the app.js file.</span></span>

5. <span data-ttu-id="556c0-109">Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `insert-paragraph`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-109">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="556c0-110">Ajoutez la fonction suivante juste après la fonction `insertParagraph` :</span><span class="sxs-lookup"><span data-stu-id="556c0-110">Just below the `insertParagraph` function, add the following function:</span></span>

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

7. <span data-ttu-id="556c0-111">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="556c0-111">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="556c0-112">Le code applique un style à un paragraphe, mais les styles peuvent également être appliqués aux plages de texte.</span><span class="sxs-lookup"><span data-stu-id="556c0-112">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="556c0-113">Appliquer un style personnalisé au texte</span><span class="sxs-lookup"><span data-stu-id="556c0-113">Apply a custom style to text</span></span>

1. <span data-ttu-id="556c0-114">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="556c0-114">Open the file index.html.</span></span>
2. <span data-ttu-id="556c0-115">En dessous de la balise `div` qui contient le bouton `apply-style`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-115">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="556c0-116">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="556c0-116">Open the app.js file.</span></span>

4. <span data-ttu-id="556c0-117">Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-style`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-117">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="556c0-118">Sous la fonction `applyStyle`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="556c0-118">Below the `applyStyle` function, add the following function:</span></span>

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

7. <span data-ttu-id="556c0-119">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="556c0-119">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="556c0-120">Le code applique un style personnalisé qui n’existe pas encore.</span><span class="sxs-lookup"><span data-stu-id="556c0-120">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="556c0-121">Vous allez créer un style nommé **MyCustomStyle** lors de l’étape [Test du complément](#test-the-add-in).</span><span class="sxs-lookup"><span data-stu-id="556c0-121">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a><span data-ttu-id="556c0-122">Modifier la police du texte</span><span class="sxs-lookup"><span data-stu-id="556c0-122">Change the font of text</span></span>

1. <span data-ttu-id="556c0-123">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="556c0-123">Open the file index.html.</span></span>
2. <span data-ttu-id="556c0-124">En dessous de la balise `div` qui contient le bouton `apply-custom-style`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-124">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="556c0-125">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="556c0-125">Open the app.js file.</span></span>

4. <span data-ttu-id="556c0-126">Sous la ligne qui attribue un gestionnaire de clics au bouton `apply-custom-style`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="556c0-126">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="556c0-127">Sous la fonction `applyCustomStyle`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="556c0-127">Below the `applyCustomStyle` function, add the following function:</span></span>

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

7. <span data-ttu-id="556c0-128">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="556c0-128">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="556c0-129">Le code obtient une référence au deuxième paragraphe en utilisant la méthode `ParagraphCollection.getFirst` chaînée à la méthode `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="556c0-129">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="556c0-130">Test du complément</span><span class="sxs-lookup"><span data-stu-id="556c0-130">Test the add-in</span></span>

1. <span data-ttu-id="556c0-131">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="556c0-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="556c0-132">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="556c0-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="556c0-133">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="556c0-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="556c0-134">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="556c0-134">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="556c0-135">Après la commande build, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="556c0-135">After the build, you restart the server.</span></span> <span data-ttu-id="556c0-136">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="556c0-136">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="556c0-137">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="556c0-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="556c0-138">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="556c0-138">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="556c0-139">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="556c0-139">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="556c0-140">Assurez-vous qu’il existe au moins trois paragraphes dans le document.</span><span class="sxs-lookup"><span data-stu-id="556c0-140">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="556c0-141">Vous pouvez sélectionner trois fois l’option **Insérer un paragraphe**.</span><span class="sxs-lookup"><span data-stu-id="556c0-141">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="556c0-142">*Vérifiez attentivement qu’aucun paragraphe vide n’apparaît à la fin du document. S’il y en a un, supprimez-le.*</span><span class="sxs-lookup"><span data-stu-id="556c0-142">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>
6. <span data-ttu-id="556c0-143">Dans Word, créez un style personnalisé nommé « MyCustomStyle ».</span><span class="sxs-lookup"><span data-stu-id="556c0-143">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="556c0-144">Vous pouvez y appliquer la mise en forme que vous souhaitez.</span><span class="sxs-lookup"><span data-stu-id="556c0-144">It can have any formatting that you want.</span></span>
7. <span data-ttu-id="556c0-145">Sélectionnez le bouton **Appliquer le style**.</span><span class="sxs-lookup"><span data-stu-id="556c0-145">Choose the **Apply Style** button.</span></span> <span data-ttu-id="556c0-146">Le style prédéfini **Référence intense** est appliqué au premier paragraphe.</span><span class="sxs-lookup"><span data-stu-id="556c0-146">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>
8. <span data-ttu-id="556c0-147">Sélectionnez le bouton **Appliquer un style personnalisé**.</span><span class="sxs-lookup"><span data-stu-id="556c0-147">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="556c0-148">Votre style personnalisé est appliqué au dernier paragraphe.</span><span class="sxs-lookup"><span data-stu-id="556c0-148">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="556c0-149">(Si rien ne semble se produire, le dernier paragraphe est peut-être vide.</span><span class="sxs-lookup"><span data-stu-id="556c0-149">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="556c0-150">Si c’est le cas, ajoutez-y du texte.)</span><span class="sxs-lookup"><span data-stu-id="556c0-150">If so, add some text to it.)</span></span>
9. <span data-ttu-id="556c0-151">Sélectionnez le bouton **Modifier la police**.</span><span class="sxs-lookup"><span data-stu-id="556c0-151">Choose the **Change Font** button.</span></span> <span data-ttu-id="556c0-152">La police Courier New, 18 pt, en gras, est appliquée au deuxième paragraphe.</span><span class="sxs-lookup"><span data-stu-id="556c0-152">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Didacticiel Word - Appliquer des styles et une police](../images/word-tutorial-apply-styles-and-font.png)
