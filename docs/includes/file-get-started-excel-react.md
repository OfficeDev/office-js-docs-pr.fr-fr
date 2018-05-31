# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="4acbe-101">Développement d’un complément Excel à l’aide de React</span><span class="sxs-lookup"><span data-stu-id="4acbe-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="4acbe-102">Dans cet article vous apprendrez à créer une macro complémentaire Excel à l’aide de React et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="4acbe-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="4acbe-103">Environnement</span><span class="sxs-lookup"><span data-stu-id="4acbe-103">Environment</span></span>

- <span data-ttu-id="4acbe-104">**Office pour ordinateur de bureau** : Assurez-vous de disposer de la dernière version d'Office.</span><span class="sxs-lookup"><span data-stu-id="4acbe-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="4acbe-105">Les commandes du complément nécessitent la version 16.0.6769.0000 ou supérieure (la version **16.0.6868.0000** est conseillée).</span><span class="sxs-lookup"><span data-stu-id="4acbe-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="4acbe-106">Apprenez à [Installer la dernière version des applications Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="4acbe-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="4acbe-107">**Office Online** : Aucune installation supplémentaire n'est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="4acbe-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="4acbe-108">Notez que la prise en charge des commandes au sein d'Office Online pour les comptes professionnels / scolaires est actuellement en préversion.</span><span class="sxs-lookup"><span data-stu-id="4acbe-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4acbe-109">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="4acbe-109">Prerequisites</span></span>

- <span data-ttu-id="4acbe-110">Installez [Create React App](https://github.com/facebookincubator/create-react-app) globalement.</span><span class="sxs-lookup"><span data-stu-id="4acbe-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="4acbe-111">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="4acbe-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="4acbe-112">Générer une nouvelle application React</span><span class="sxs-lookup"><span data-stu-id="4acbe-112">Generate a new React app</span></span>

<span data-ttu-id="4acbe-113">L’outil Create React App permet de générer votre application React.</span><span class="sxs-lookup"><span data-stu-id="4acbe-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="4acbe-114">À partir du terminal, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="4acbe-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="4acbe-115">Générer le fichier manifeste et charger une version test du complément</span><span class="sxs-lookup"><span data-stu-id="4acbe-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="4acbe-116">Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="4acbe-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="4acbe-117">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="4acbe-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="4acbe-118">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="4acbe-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="4acbe-119">Exécutez la commande suivante, puis répondez aux invites comme indiqué dans la capture d’écran suivante :</span><span class="sxs-lookup"><span data-stu-id="4acbe-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="4acbe-120">**Voulez-vous créer un sous-dossier de votre projet ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="4acbe-120">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="4acbe-121">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="4acbe-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="4acbe-122">**Quelle application client Office voulez-vous prendre en charge ? :** `Excel`</span><span class="sxs-lookup"><span data-stu-id="4acbe-122">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="4acbe-123">**Voulez-vous créer un complément ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="4acbe-123">**Would you like to create a new add-in?:** `No`</span></span>

    <span data-ttu-id="4acbe-p105">Le générateur demande ensuite si vous voulez ouvrir **resource.html**. Il n’est pas nécessaire de l’ouvrir pour ce didacticiel, mais n’hésitez pas à l’ouvrir si vous êtes curieux. Cliquez sur Oui ou Non pour fermer l’assistant et laisser le générateur faire son travail.</span><span class="sxs-lookup"><span data-stu-id="4acbe-p105">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Générateur Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="4acbe-128">Si vous êtes invité à remplacer **package.json**, répondez **Non** (ne pas remplacer).</span><span class="sxs-lookup"><span data-stu-id="4acbe-128">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="4acbe-129">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="4acbe-129">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="4acbe-130">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="4acbe-130">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="4acbe-131">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="4acbe-131">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="4acbe-132">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="4acbe-132">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="4acbe-133">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="4acbe-133">Update the app</span></span>

1. <span data-ttu-id="4acbe-134">Ouvrez **public/index.html**, ajoutez la balise `<script>` suivante immédiatement avant la balise `</head>`, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4acbe-134">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="4acbe-135">Ouvrez **src/index.js**, remplacez `ReactDOM.render(<App />, document.getElementById('root'));` par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4acbe-135">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="4acbe-136">Ouvrez **src/App.js**, remplacez le contenu du fichier par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4acbe-136">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. <span data-ttu-id="4acbe-137">Ouvrez **src/App.css**, remplacez le contenu du fichier par le code CSS suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4acbe-137">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="try-it-out"></a><span data-ttu-id="4acbe-138">Essayez !</span><span class="sxs-lookup"><span data-stu-id="4acbe-138">Try it out</span></span>

1. <span data-ttu-id="4acbe-139">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="4acbe-139">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="4acbe-140">Windows :</span><span class="sxs-lookup"><span data-stu-id="4acbe-140">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="4acbe-141">macOS :</span><span class="sxs-lookup"><span data-stu-id="4acbe-141">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="4acbe-p106">Une fenêtre de navigateur s’ouvre avec le complément qu’elle contient. Fermez cette fenêtre.</span><span class="sxs-lookup"><span data-stu-id="4acbe-p106">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="4acbe-144">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="4acbe-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="4acbe-146">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="4acbe-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="4acbe-147">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="4acbe-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="4acbe-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="4acbe-149">Next steps</span></span>

<span data-ttu-id="4acbe-p107">Félicitations, vous avez créé un complément Excel à l’aide de React ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="4acbe-p107">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="4acbe-152">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="4acbe-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="4acbe-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4acbe-153">See also</span></span>

* [<span data-ttu-id="4acbe-154">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="4acbe-154">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="4acbe-155">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4acbe-155">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="4acbe-156">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="4acbe-156">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="4acbe-157">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4acbe-157">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
