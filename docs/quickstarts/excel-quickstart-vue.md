---
title: Créer un complément de volet de tâches Excel à l’aide de Vue
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et de Vue pour Office.
ms.date: 06/16/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: cd709910c9e69478c953c03b5e17d5512e875d91
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007817"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="5d711-103">Créer un complément de volet de tâches Excel à l’aide de Vue</span><span class="sxs-lookup"><span data-stu-id="5d711-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="5d711-104">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de Vue et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="5d711-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5d711-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="5d711-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="5d711-106">Installez l’[interface de ligne de commande Vue](https://cli.vuejs.org/) globalement.</span><span class="sxs-lookup"><span data-stu-id="5d711-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="5d711-107">Génération d’une nouvelle application Vue</span><span class="sxs-lookup"><span data-stu-id="5d711-107">Generate a new Vue app</span></span>

<span data-ttu-id="5d711-p101">Utilisez l’interface de ligne de commande Vue pour générer une nouvelle application Vue. À partir du terminal, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="5d711-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="5d711-110">Sélectionnez ensuite la `Default` prédéfinie pour « Vue 3 » (vous pouvez choisir d’utiliser « Vue 2 » si vous préférez).</span><span class="sxs-lookup"><span data-stu-id="5d711-110">Then select the `Default` preset for "Vue 3" (you may choose to use "Vue 2" if you'd prefer).</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="5d711-111">Génération du fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="5d711-111">Generate the manifest file</span></span>

<span data-ttu-id="5d711-112">Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="5d711-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="5d711-113">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="5d711-113">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="5d711-114">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément en exécutant la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="5d711-114">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="5d711-115">Lorsque vous exécutez la commande `yo office`, il est possible que vous receviez des messages d’invite sur les règles de collecte de données de Yeoman et les outils CLI de complément Office.</span><span class="sxs-lookup"><span data-stu-id="5d711-115">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="5d711-116">Utilisez les informations fournies pour répondre aux invites comme vous l’entendez.</span><span class="sxs-lookup"><span data-stu-id="5d711-116">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="5d711-117">Si vous sélectionnez **Quitter** en réponse à la deuxième invite, vous devez réexécuter la commande `yo office` lorsque vous êtes prêt à créer votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="5d711-117">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="5d711-118">Lorsque vous y êtes invité, fournissez les informations suivantes pour créer votre projet de complément :</span><span class="sxs-lookup"><span data-stu-id="5d711-118">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="5d711-119">**Sélectionnez un type de projet :** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="5d711-119">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="5d711-120">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="5d711-120">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="5d711-121">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="5d711-121">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Capture d’écran des invites d’interface de ligne de commande du générateur de compléments Yeoman Office pour les projets de fonctions personnalisées](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="5d711-123">Une fois que vous avez terminé les étapes de l’Assistant, celui-ci crée un dossier `My Office Add-in` qui contient un fichier `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="5d711-123">After you complete the wizard, it creates a `My Office Add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="5d711-124">Vous utiliserez le manifeste pour charger une version test et tester votre complément à la fin du Démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="5d711-124">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="5d711-125">Vous pouvez ignorer les *instructions suivantes* fournies par le générateur Yeoman une fois que le complément a été créé.</span><span class="sxs-lookup"><span data-stu-id="5d711-125">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="5d711-126">Les instructions détaillées de cet article fournissent tous les conseils nécessaires à l’exécution de ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="5d711-126">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="5d711-127">Sécurisation de l’application</span><span class="sxs-lookup"><span data-stu-id="5d711-127">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. <span data-ttu-id="5d711-128">Pour activer HTTPS pour votre application, créez un fichier `vue.config.js` dans le dossier racine du projet Vue avec le contenu suivant :</span><span class="sxs-lookup"><span data-stu-id="5d711-128">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: true,
        key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
        cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
        ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`))
      }
    }
    ```

2. <span data-ttu-id="5d711-129">À partir du terminal, exécutez la commande suivante pour installer les certificats du complément.</span><span class="sxs-lookup"><span data-stu-id="5d711-129">From the terminal, run the following command to install the add-in's certificates.</span></span>

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="update-the-app"></a><span data-ttu-id="5d711-130">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="5d711-130">Update the app</span></span>

1. <span data-ttu-id="5d711-131">Ouvrez le fichier `public/index.html` et ajoutez la balise `<script>` suivante juste avant la balise `</head>` :</span><span class="sxs-lookup"><span data-stu-id="5d711-131">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="5d711-132">Ouvrez `src/main.js` et remplacez le contenu par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="5d711-132">Open `src/main.js` and replace the contents with the following code:</span></span>

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

3. <span data-ttu-id="5d711-133">Ouvrez `src/App.vue` et remplacez le contenu du fichier par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="5d711-133">Open `src/App.vue` and replace the file contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div class="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
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
   </style>
   ```

## <a name="start-the-dev-server"></a><span data-ttu-id="5d711-134">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="5d711-134">Start the dev server</span></span>

1. <span data-ttu-id="5d711-135">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="5d711-135">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="5d711-136">Dans un navigateur web, accédez à `https://localhost:3000` (remarquez le `https`).</span><span class="sxs-lookup"><span data-stu-id="5d711-136">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="5d711-137">Si la page sur `https://localhost:3000` est vide et qu’aucune erreur de certificat ne s’affiche, cela signifie qu’elle fonctionne.</span><span class="sxs-lookup"><span data-stu-id="5d711-137">If the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="5d711-138">L’application Vue est montée une fois qu’Office est initialisé, de sorte qu’elle affiche uniquement les éléments dans un environnement Excel.</span><span class="sxs-lookup"><span data-stu-id="5d711-138">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="5d711-139">Essayez</span><span class="sxs-lookup"><span data-stu-id="5d711-139">Try it out</span></span>

1. <span data-ttu-id="5d711-140">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5d711-140">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="5d711-141">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="5d711-141">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="5d711-142">Navigateur web : [Chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="5d711-142">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="5d711-143">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="5d711-143">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="5d711-144">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="5d711-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Capture d’écran du menu Accueil d’Excel, avec le bouton Afficher le volet Office mis en évidence](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="5d711-146">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5d711-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5d711-147">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="5d711-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Capture d’écran d’Excel avec le volet Office Complément ouvert](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="5d711-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="5d711-149">Next steps</span></span>

<span data-ttu-id="5d711-p106">Félicitations, vous avez créé un complément du volet Office Excel à l’aide de Vue ! Maintenant, apprenez-en davantage sur les fonctionnalités d’un complément Excel et créez un complément plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="5d711-p106">Congratulations, you've successfully created an Excel task pane add-in using Vue! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5d711-152">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5d711-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="5d711-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5d711-153">See also</span></span>

* [<span data-ttu-id="5d711-154">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="5d711-154">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="5d711-155">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="5d711-155">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="5d711-156">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="5d711-156">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5d711-157">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5d711-157">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="5d711-158">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="5d711-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
