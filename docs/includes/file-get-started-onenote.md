# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="4d0bb-101">Cr?er votre premier compl?ment OneNote</span><span class="sxs-lookup"><span data-stu-id="4d0bb-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="4d0bb-102">Cet article d?crit le processus de cr?ation d?un compl?ment OneNote ? l?aide de jQuery et de l?API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4d0bb-103">Conditions pr?alables</span><span class="sxs-lookup"><span data-stu-id="4d0bb-103">Prerequisites</span></span>

- [<span data-ttu-id="4d0bb-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="4d0bb-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="4d0bb-105">Installez la derni?re version de [Yeoman](https://github.com/yeoman/yo) et le [g?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="4d0bb-106">Cr?ation du projet de compl?ment</span><span class="sxs-lookup"><span data-stu-id="4d0bb-106">Create the add-in project</span></span>

1. <span data-ttu-id="4d0bb-107">Cr?ez un dossier sur votre lecteur local et nommez-le `my-onenote-addin`.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="4d0bb-108">Il s?agit de l?emplacement dans lequel vous allez cr?er les fichiers de votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="4d0bb-109">Acc?dez ? votre nouveau dossier.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="4d0bb-110">Utilisez le g?n?rateur Yeoman afin de cr?er un projet de compl?ment OneNote.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="4d0bb-111">Ex?cutez la commande suivante, puis r?pondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="4d0bb-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="4d0bb-112">**Voulez-vous cr?er un sous-dossier de votre projet ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-112">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="4d0bb-113">**Comment souhaitez-vous nommer votre compl?ment ? :** `OneNote Add-in`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-113">**What do you want to name your add-in?:** `OneNote Add-in`</span></span>
    - <span data-ttu-id="4d0bb-114">**Quelle application client Office voulez-vous prendre en charge ? :** `OneNote`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-114">**Which Office client application would you like to support?:** `OneNote`</span></span>
    - <span data-ttu-id="4d0bb-115">**Voulez-vous cr?er un compl?ment ? :** `Yes`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-115">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="4d0bb-116">**Souhaitez-vous utiliser TypeScript ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-116">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="4d0bb-117">**Choisissez une infrastructure :** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="4d0bb-117">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="4d0bb-p103">Le g?n?rateur demande ensuite si vous voulez ouvrir **resource.html**. Il n?est pas n?cessaire de l?ouvrir pour ce didacticiel, mais n?h?sitez pas ? l?ouvrir si vous ?tes curieux. Cliquez sur Oui ou Non pour fermer l?assistant et laisser le g?n?rateur faire son travail.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Capture d??cran des invites et des r?ponses relatives au g?n?rateur Yeoman](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a><span data-ttu-id="4d0bb-122">Mise ? jour du code</span><span class="sxs-lookup"><span data-stu-id="4d0bb-122">Update the code</span></span>

1. <span data-ttu-id="4d0bb-123">Dans votre ?diteur de code, ouvrez **index.html** ? la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-123">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="4d0bb-124">Ce fichier contient le code HTML qui s?affichera dans le volet Office du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-124">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="4d0bb-125">Remplacez l??l?ment `<main>` dans l??l?ment `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-125">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="4d0bb-126">Cette option ajoute une zone de texte et un bouton ? l?aide des [composants de la structure de l?interface utilisateur d?Office](http://dev.office.com/fabric/components).</span><span class="sxs-lookup"><span data-stu-id="4d0bb-126">This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).</span></span>

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. <span data-ttu-id="4d0bb-127">Ouvrez le fichier **app.js** pour sp?cifier le script pour le compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-127">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="4d0bb-128">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-128">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="4d0bb-129">Mise ? jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="4d0bb-129">Update the manifest</span></span>

1. <span data-ttu-id="4d0bb-130">Ouvrez le fichier nomm? **one-note-add-in-manifest.xml** pour d?finir les param?tres et les fonctionnalit?s du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-130">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="4d0bb-131">L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-131">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="4d0bb-132">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-132">Replace it with your name.</span></span>

3. <span data-ttu-id="4d0bb-133">L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-133">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="4d0bb-134">Remplacez-le par **A task pane add-in for OneNote**.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-134">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="4d0bb-135">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-135">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="4d0bb-136">D?marrage du serveur de d?veloppement</span><span class="sxs-lookup"><span data-stu-id="4d0bb-136">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="4d0bb-137">Essayez !</span><span class="sxs-lookup"><span data-stu-id="4d0bb-137">Try it out</span></span>

1. <span data-ttu-id="4d0bb-138">Dans [OneNote Online](https://www.onenote.com/notebooks), ouvrez un bloc-notes.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-138">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="4d0bb-139">Choisissez **Insertion > Compl?ments Office** pour ouvrir la bo?te de dialogue Compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-139">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="4d0bb-140">Si vous ?tes connect? avec votre compte de consommateur, s?lectionnez l?onglet **MES COMPL?MENTS**, puis choisissez **T?l?charger mon compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-140">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="4d0bb-141">Si vous ?tes connect? avec votre compte professionnel ou scolaire, s?lectionnez l?onglet **MON ORGANISATION**, puis choisissez **T?l?charger mon compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-141">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="4d0bb-142">L?image suivante montre l?onglet **MES COMPL?MENTS** pour les blocs-notes de consommateurs.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-142">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="4d0bb-143">Dans la bo?te de dialogue T?l?charger le compl?ment, acc?dez ? **one-note-add-in-manifest.xml** dans le dossier de projet, puis choisissez **T?l?charger**.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-143">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="4d0bb-144">Depuis l?onglet **Accueil**, cliquez le bouton **Afficher le volet Office** du ruban.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-144">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="4d0bb-145">Le compl?ment volet Office s?ouvre dans un iFrame ? c?t? de la page OneNote.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-145">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="4d0bb-146">Entrez du texte dans la zone de texte, puis choisissez **Ajouter un plan**.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-146">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="4d0bb-147">Le texte que vous avez entr? est ajout? ? la page.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-147">The text you entered is added to the page.</span></span> 

    ![Compl?ment OneNote g?n?r? ? partir de cette proc?dure pas ? pas](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="4d0bb-149">Conseils et r?solution des probl?mes</span><span class="sxs-lookup"><span data-stu-id="4d0bb-149">Troubleshooting and tips</span></span>

- <span data-ttu-id="4d0bb-p111">Vous pouvez d?boguer le compl?ment ? l?aide des outils de d?veloppement de votre navigateur. Lorsque vous utilisez le serveur web Gulp et le d?bogage dans Internet Explorer ou Chrome, vous pouvez enregistrer les modifications localement et simplement actualiser l?iFrame du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-p111">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="4d0bb-p112">Lorsque vous examinez un objet OneNote, les propri?t?s qui sont actuellement disponibles affichent les valeurs r?elles. Les propri?t?s qui doivent ?tre charg?es sont affich?es comme *non d?finies*. D?veloppez le n?ud `_proto_` pour visualiser les propri?t?s qui sont d?finies sur l?objet, mais qui ne sont pas encore charg?es.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-p112">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Objet OneNote d?charg? dans le d?bogueur](../images/onenote-debug.png)

- <span data-ttu-id="4d0bb-p113">Vous devez activer le contenu mixte dans le navigateur si votre compl?ment utilise des ressources HTTP. Les compl?ments de production doivent uniquement utiliser des ressources HTTPS s?curis?es.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-p113">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="4d0bb-158">Les compl?ments de volet Office peuvent ?tre ouverts ? partir de n?importe o?, mais les compl?ments de contenu peuvent uniquement ?tre ins?r?s ? l?int?rieur de contenu de page normal (et non dans des titres, des images, des iFrames, etc.).</span><span class="sxs-lookup"><span data-stu-id="4d0bb-158">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="4d0bb-159">?tapes suivantes</span><span class="sxs-lookup"><span data-stu-id="4d0bb-159">Next steps</span></span>

<span data-ttu-id="4d0bb-160">F?licitations, vous avez cr?? un compl?ment OneNote !</span><span class="sxs-lookup"><span data-stu-id="4d0bb-160">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="4d0bb-161">Ensuite, vous allez ?tudier en d?tail les concepts fondamentaux de la cr?ation de compl?ments Excel.</span><span class="sxs-lookup"><span data-stu-id="4d0bb-161">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="4d0bb-162">Vue d?ensemble de la programmation de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="4d0bb-162">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="4d0bb-163">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4d0bb-163">See also</span></span>

- [<span data-ttu-id="4d0bb-164">Vue d?ensemble de la programmation de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="4d0bb-164">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="4d0bb-165">R?f?rence de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="4d0bb-165">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="4d0bb-166">Exemple de grille d??valuation</span><span class="sxs-lookup"><span data-stu-id="4d0bb-166">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="4d0bb-167">Vue d?ensemble de la plateforme des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="4d0bb-167">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
