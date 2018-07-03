# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="3c2df-101">Créer un complément Excel à l’aide d’Angular</span><span class="sxs-lookup"><span data-stu-id="3c2df-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="3c2df-102">Dans cet article, vous allez découvrir le processus de création d’un complément Excel à l’aide d’Angular et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="3c2df-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3c2df-103">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="3c2df-103">Prerequisites</span></span>

- <span data-ttu-id="3c2df-104">Assurez-vous que vous avez la [configuration requise pour CLI Angular](https://github.com/angular/angular-cli#prerequisites) et installez tous les composants manquants.</span><span class="sxs-lookup"><span data-stu-id="3c2df-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="3c2df-105">Installez [CLI Angular](https://github.com/angular/angular-cli) globalement.</span><span class="sxs-lookup"><span data-stu-id="3c2df-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="3c2df-106">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="3c2df-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="3c2df-107">Générer une nouvelle application Angular</span><span class="sxs-lookup"><span data-stu-id="3c2df-107">Generate a new Angular app</span></span>

<span data-ttu-id="3c2df-108">Utilisez l’outil CLI Angular pour générer votre application Angular.</span><span class="sxs-lookup"><span data-stu-id="3c2df-108">Use the Angular CLI to generate your Angular app.</span></span> <span data-ttu-id="3c2df-109">À partir du terminal, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="3c2df-109">From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="3c2df-110">Génération du fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="3c2df-110">Generate the manifest file</span></span>

<span data-ttu-id="3c2df-111">Le fichier manifeste d’un complément définit ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="3c2df-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="3c2df-112">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="3c2df-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="3c2df-113">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="3c2df-113">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="3c2df-114">Exécutez la commande suivante, puis répondez aux invites comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="3c2df-114">Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="3c2df-115">**Choisissez un type de projet :** `Manifest`</span><span class="sxs-lookup"><span data-stu-id="3c2df-115">**Choose a project type:** `Manifest`</span></span>
    - <span data-ttu-id="3c2df-116">**Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="3c2df-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="3c2df-117">**Quelle application client Office voulez-vous prendre en charge ?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="3c2df-117">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="3c2df-118">Après avoir terminé l'assistant, un fichier manifeste et un fichier de ressources sont disponibles pour vous permettre de générer votre projet.</span><span class="sxs-lookup"><span data-stu-id="3c2df-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Générateur Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="3c2df-120">Si vous êtes invité à remplacer **package.json**, répondez **Non** (ne pas remplacer).</span><span class="sxs-lookup"><span data-stu-id="3c2df-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="3c2df-121">Sécurisation de l’application</span><span class="sxs-lookup"><span data-stu-id="3c2df-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="3c2df-122">Pour ce démarrage rapide, vous pouvez utiliser les certificats fournis par le **générateur de compléments Office Yeoman**.</span><span class="sxs-lookup"><span data-stu-id="3c2df-122">For this quickstart, you can use the certificates that the **Yeoman generator for Office Add-ins** provides.</span></span> <span data-ttu-id="3c2df-123">Vous avez déjà installé le générateur globalement (comme demandé dans la section **Conditions préalables** de ce démarrage rapide). Vous n’avez donc qu’à copier les certificats situés dans l’emplacement d’installation global dans le dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="3c2df-123">You've already installed the generator globally (as part of the **Prerequisites** for this quickstart), so you'll just need to copy the certificates from the global install location into your app folder.</span></span> <span data-ttu-id="3c2df-124">La procédure suivante explique comment effectuer cette procédure.</span><span class="sxs-lookup"><span data-stu-id="3c2df-124">The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="3c2df-125">À partir du terminal, exécutez la commande suivante pour identifier le dossier où les bibliothèques **npm** globales sont installées :</span><span class="sxs-lookup"><span data-stu-id="3c2df-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="3c2df-126">La première ligne de la sortie générée par cette commande spécifie le dossier où les bibliothèques **npm** globales sont installées.</span><span class="sxs-lookup"><span data-stu-id="3c2df-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="3c2df-127">À l’aide de l’explorateur de fichiers, accédez au dossier `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`.</span><span class="sxs-lookup"><span data-stu-id="3c2df-127">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder.</span></span> <span data-ttu-id="3c2df-128">À partir de cet emplacement, copiez le dossier `certs` dans votre presse-papiers.</span><span class="sxs-lookup"><span data-stu-id="3c2df-128">From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="3c2df-129">Accédez au dossier racine de l’application Angular que vous avez créée à l’étape 1 de la section précédente et collez le dossier `certs` (qui se trouve dans votre presse-papiers) dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="3c2df-130">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="3c2df-130">Update the app</span></span>

1. <span data-ttu-id="3c2df-131">Dans votre éditeur de code, ouvrez **package.json** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="3c2df-131">In your code editor, open **package.json** in the root of the project.</span></span> <span data-ttu-id="3c2df-132">Modifiez le script `start` pour spécifier que le serveur doit s’exécuter à l’aide de SSL et du port 3000, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-132">Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="3c2df-133">Ouvrez **.angular-cli.json** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="3c2df-133">Open **.angular-cli.json** in the root of the project.</span></span> <span data-ttu-id="3c2df-134">Modifiez l’objet **defaults** pour indiquer l’emplacement des fichiers de certificat et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-134">Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. <span data-ttu-id="3c2df-135">Ouvrez **src/index.html**, ajoutez la balise `<script>` suivante immédiatement avant la balise `</head>`, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="3c2df-136">Ouvrez **src/main.ts**, remplacez `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="3c2df-137">Ouvrez **src/polyfills.ts**, ajoutez la ligne de code suivante au-dessus de toutes les instructions `import` existantes, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="3c2df-138">Dans **src/polyfills.ts**, supprimez les commentaires des lignes suivantes, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. <span data-ttu-id="3c2df-139">Ouvrez **src/app/app.component.html**, remplacez le contenu du fichier par le code HTML suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. <span data-ttu-id="3c2df-140">Ouvrez **src/app/app.component.css**, remplacez le contenu du fichier par le code CSS suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="3c2df-141">Ouvrez **src/app/app.component.ts**, remplacez le contenu du fichier par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="3c2df-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="3c2df-142">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="3c2df-142">Start the dev server</span></span>

1. <span data-ttu-id="3c2df-143">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="3c2df-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="3c2df-p107">Dans un navigateur web, accédez à `https://localhost:3000`. Si votre navigateur indique que le certificat du site n’est pas approuvé, vous devrez ajouter le certificat en tant que certificat approuvé. Consultez la rubrique relative à l’[ajout de certificats auto-signés en tant que certificats racine approuvés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour obtenir plus de détails.</span><span class="sxs-lookup"><span data-stu-id="3c2df-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3c2df-147">Il est possible que le navigateur web Chrome continue d’indiquer que le certificat du site n’est pas approuvé, même si vous avez suivi les étapes décrites dans l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="3c2df-147">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="3c2df-148">Vous pouvez ignorer ce message d’avertissement dans Chrome. Vérifiez tout de même que le certificat est approuvé en entrant `https://localhost:3000` dans Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="3c2df-148">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="3c2df-149">Une fois que votre navigateur a chargé la page du complément sans erreurs de certificat, vous pouvez tester votre complément.</span><span class="sxs-lookup"><span data-stu-id="3c2df-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="3c2df-150">Essayez !</span><span class="sxs-lookup"><span data-stu-id="3c2df-150">Try it out</span></span>

1. <span data-ttu-id="3c2df-151">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="3c2df-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="3c2df-152">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3c2df-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="3c2df-153">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="3c2df-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="3c2df-154">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3c2df-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="3c2df-155">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="3c2df-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="3c2df-157">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3c2df-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="3c2df-158">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="3c2df-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="3c2df-160">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="3c2df-160">Next steps</span></span>

<span data-ttu-id="3c2df-p109">Félicitations, vous avez créé un complément Excel à l’aide d’Angular ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="3c2df-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3c2df-163">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="3c2df-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="3c2df-164">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3c2df-164">See also</span></span>

* [<span data-ttu-id="3c2df-165">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="3c2df-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="3c2df-166">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="3c2df-166">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="3c2df-167">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="3c2df-167">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="3c2df-168">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="3c2df-168">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
