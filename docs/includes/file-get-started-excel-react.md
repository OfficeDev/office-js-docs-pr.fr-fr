# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="5d91a-101">Développement d’un complément Excel à l’aide de React</span><span class="sxs-lookup"><span data-stu-id="5d91a-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="5d91a-102">Cet article décrit le processus de création d’un complément Excel à l’aide de React et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="5d91a-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="5d91a-103">Environnement</span><span class="sxs-lookup"><span data-stu-id="5d91a-103">Environment</span></span>

- <span data-ttu-id="5d91a-p101">**Office pour ordinateur de bureau**:  Assurez-vous de disposer de la dernière version d'Office. Les commandes du complément nécessitent la version 16.0.6769.0000 ou supérieure (la version**16.0.6868.0000** est conseillée). Apprenez à [Installer la dernière version des applications Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="5d91a-p101">**Office Desktop**: Ensure that you have the latest version of Office installed. Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended). Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="5d91a-p102">**Office Online**: Aucune installation supplémentaire n'est nécessaire. Notez que la prise en charge des commandes au sein d'Office Online pour les comptes professionnels / scolaires est actuellement en préversion.</span><span class="sxs-lookup"><span data-stu-id="5d91a-p102">**Office Online**: There is no additional setup. Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5d91a-109">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="5d91a-109">Prerequisites</span></span>

- [<span data-ttu-id="5d91a-110">Node.js</span><span class="sxs-lookup"><span data-stu-id="5d91a-110">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="5d91a-111">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="5d91a-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="5d91a-112">Création de l’application web</span><span class="sxs-lookup"><span data-stu-id="5d91a-112">Create the web app</span></span>

1. <span data-ttu-id="5d91a-p103">Créez un dossier sur votre lecteur local et nommez-le **my-addin**. Il s’agit de l’endroit où vous allez créer les fichiers de votre application.</span><span class="sxs-lookup"><span data-stu-id="5d91a-p103">Create a folder on your local drive and name it **my-addin**. This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="5d91a-115">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="5d91a-115">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="5d91a-p104">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="5d91a-p104">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown in the following screenshot.</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="5d91a-118">**Choisissez un type de projet :** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="5d91a-118">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="5d91a-119">**Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="5d91a-119">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="5d91a-120">**Quelle application client Office voulez-vous prendre en charge ? :** `Excel`</span><span class="sxs-lookup"><span data-stu-id="5d91a-120">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Générateur Yeoman](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="5d91a-122">Une fois que vous avez terminé avec l'assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="5d91a-122">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

4.  <span data-ttu-id="5d91a-123">Ouvrez **src/components/App.tsx**, recherchez le commentaire « Mettre à jour la couleur de remplissage », puis modifiez la couleur de remplissage de « jaune » à « bleu » avant d'enregistrer le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d91a-123">Open **src/components/App.tsx**, search for the comment "Update the fill color," then change the fill color from 'yellow' to 'blue', and save the file.</span></span> 

    ```js
    range.format.fill.color = 'blue'

    ```

5. <span data-ttu-id="5d91a-124">Dans le bloc `return` de la fonction `render` au sein de **src/components/App.tsx**, mettez le `<Herolist>` à jour avec le code ci-dessous, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d91a-124">In the `return` block of the `render` function within **src/components/App.tsx**, update the `<Herolist>` to the code below, and save the file.</span></span> 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. <span data-ttu-id="5d91a-125">Effectuez les étapes décrites dans la rubrique relative à l’[Ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="5d91a-125">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

7. <span data-ttu-id="5d91a-p105">Chargez une version test de votre complément afin qu’il apparaisse dans Excel. Dans le terminal, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="5d91a-p105">Sideload your add-in so it will appear in Excel. In the terminal, run the following command:</span></span> 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a><span data-ttu-id="5d91a-128">Essayez</span><span class="sxs-lookup"><span data-stu-id="5d91a-128">Try it out</span></span>

1. <span data-ttu-id="5d91a-129">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="5d91a-129">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="5d91a-130">Windows :</span><span class="sxs-lookup"><span data-stu-id="5d91a-130">Windows:</span></span>
    ```bash
    npm start
    ```

2. <span data-ttu-id="5d91a-131">Dans Word, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="5d91a-131">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton de Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="5d91a-133">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5d91a-133">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5d91a-134">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en bleu.</span><span class="sxs-lookup"><span data-stu-id="5d91a-134">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="5d91a-136">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="5d91a-136">Next steps</span></span>

<span data-ttu-id="5d91a-p106">Félicitations, vous avez créé un complément Excel à l’aide de React ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="5d91a-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5d91a-139">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5d91a-139">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="5d91a-140">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5d91a-140">See also</span></span>

* [<span data-ttu-id="5d91a-141">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5d91a-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="5d91a-142">Concepts fondamentaux de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="5d91a-142">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5d91a-143">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5d91a-143">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="5d91a-144">Référence de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="5d91a-144">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
