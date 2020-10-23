---
title: Créer un complément de volet de tâches Excel à l’aide de Vue
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et de Vue pour Office.
ms.date: 10/14/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aff6271fa4d602141807b33ff96637957818c466
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741168"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="daf0b-103">Créer un complément de volet de tâches Excel à l’aide de Vue</span><span class="sxs-lookup"><span data-stu-id="daf0b-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="daf0b-104">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de Vue et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="daf0b-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="daf0b-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="daf0b-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="daf0b-106">Installez l’[interface de ligne de commande Vue](https://cli.vuejs.org/) globalement.</span><span class="sxs-lookup"><span data-stu-id="daf0b-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="daf0b-107">Génération d’une nouvelle application Vue</span><span class="sxs-lookup"><span data-stu-id="daf0b-107">Generate a new Vue app</span></span>

<span data-ttu-id="daf0b-p101">Utilisez l’interface de ligne de commande Vue pour générer une nouvelle application Vue. À partir du terminal, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="daf0b-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="daf0b-110">Ensuite, sélectionnez la présélection `default`.</span><span class="sxs-lookup"><span data-stu-id="daf0b-110">Then select the `default` preset.</span></span> <span data-ttu-id="daf0b-111">Si vous êtes invité à utiliser Yarn ou NPM comme package, vous pouvez choisir l’un ou l’autre.</span><span class="sxs-lookup"><span data-stu-id="daf0b-111">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="daf0b-112">Génération du fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="daf0b-112">Generate the manifest file</span></span>

<span data-ttu-id="daf0b-113">Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="daf0b-113">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="daf0b-114">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="daf0b-114">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="daf0b-115">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément en exécutant la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="daf0b-115">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="daf0b-116">Lorsque vous exécutez la commande `yo office`, il est possible que vous receviez des messages d’invite sur les règles de collecte de données de Yeoman et les outils CLI de complément Office.</span><span class="sxs-lookup"><span data-stu-id="daf0b-116">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="daf0b-117">Utilisez les informations fournies pour répondre aux invites comme vous l’entendez.</span><span class="sxs-lookup"><span data-stu-id="daf0b-117">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="daf0b-118">Si vous sélectionnez **Quitter** en réponse à la deuxième invite, vous devez réexécuter la commande `yo office` lorsque vous êtes prêt à créer votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="daf0b-118">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="daf0b-119">Lorsque vous y êtes invité, fournissez les informations suivantes pour créer votre projet de complément :</span><span class="sxs-lookup"><span data-stu-id="daf0b-119">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="daf0b-120">**Sélectionnez un type de projet :** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="daf0b-120">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="daf0b-121">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="daf0b-121">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="daf0b-122">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="daf0b-122">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Générateur Yeoman](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="daf0b-124">Une fois que vous avez terminé les étapes de l’Assistant, celui-ci crée un dossier `My Office Add-in` qui contient un fichier `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="daf0b-124">After you complete the wizard, it creates a `My Office Add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="daf0b-125">Vous utiliserez le manifeste pour charger une version test et tester votre complément à la fin du Démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="daf0b-125">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="daf0b-126">Vous pouvez ignorer les *instructions suivantes* fournies par le générateur Yeoman une fois que le complément a été créé.</span><span class="sxs-lookup"><span data-stu-id="daf0b-126">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="daf0b-127">Les instructions détaillées de cet article fournissent tous les conseils nécessaires à l’exécution de ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="daf0b-127">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="daf0b-128">Sécurisation de l’application</span><span class="sxs-lookup"><span data-stu-id="daf0b-128">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. <span data-ttu-id="daf0b-129">Pour activer HTTPS pour votre application, créez un fichier `vue.config.js` dans le dossier racine du projet Vue avec le contenu suivant :</span><span class="sxs-lookup"><span data-stu-id="daf0b-129">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

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

2. <span data-ttu-id="daf0b-130">À partir du terminal, exécutez la commande suivante pour installer les certificats du complément.</span><span class="sxs-lookup"><span data-stu-id="daf0b-130">From the terminal, run the following command to install the add-in's certificates.</span></span>

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="update-the-app"></a><span data-ttu-id="daf0b-131">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="daf0b-131">Update the app</span></span>

1. <span data-ttu-id="daf0b-132">Ouvrez le fichier `public/index.html` et ajoutez la balise `<script>` suivante juste avant la balise `</head>` :</span><span class="sxs-lookup"><span data-stu-id="daf0b-132">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="daf0b-133">Ouvrez `src/main.js` et remplacez le contenu par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="daf0b-133">Open `src/main.js` and replace the contents with the following code:</span></span>

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. <span data-ttu-id="daf0b-134">Ouvrez `src/App.vue` et remplacez le contenu du fichier par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="daf0b-134">Open `src/App.vue` and replace the file contents with the following code:</span></span>

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

## <a name="start-the-dev-server"></a><span data-ttu-id="daf0b-135">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="daf0b-135">Start the dev server</span></span>

1. <span data-ttu-id="daf0b-136">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="daf0b-136">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="daf0b-137">Dans un navigateur web, accédez à `https://localhost:3000` (remarquez le `https`).</span><span class="sxs-lookup"><span data-stu-id="daf0b-137">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="daf0b-138">Si la page sur `https://localhost:3000` est vide et qu’aucune erreur de certificat ne s’affiche, cela signifie qu’elle fonctionne.</span><span class="sxs-lookup"><span data-stu-id="daf0b-138">If the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="daf0b-139">L’application Vue est montée une fois qu’Office est initialisé, de sorte qu’elle affiche uniquement les éléments dans un environnement Excel.</span><span class="sxs-lookup"><span data-stu-id="daf0b-139">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="daf0b-140">Essayez</span><span class="sxs-lookup"><span data-stu-id="daf0b-140">Try it out</span></span>

1. <span data-ttu-id="daf0b-141">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="daf0b-141">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="daf0b-142">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="daf0b-142">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="daf0b-143">Navigateur web : [Chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="daf0b-143">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="daf0b-144">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="daf0b-144">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="daf0b-145">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="daf0b-145">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="daf0b-147">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="daf0b-147">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="daf0b-148">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="daf0b-148">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="daf0b-150">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="daf0b-150">Next steps</span></span>

<span data-ttu-id="daf0b-151">Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide de Vue !</span><span class="sxs-lookup"><span data-stu-id="daf0b-151">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="daf0b-152">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="daf0b-152">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="daf0b-153">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="daf0b-153">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="daf0b-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="daf0b-154">See also</span></span>

* [<span data-ttu-id="daf0b-155">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="daf0b-155">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="daf0b-156">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="daf0b-156">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="daf0b-157">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="daf0b-157">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="daf0b-158">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="daf0b-158">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="daf0b-159">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="daf0b-159">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
