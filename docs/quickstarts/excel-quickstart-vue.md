---
title: Créer un complément de volet de tâches Excel à l’aide de Vue
description: ''
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d3c759579ebd19cc1f53f68d69db04768b69e1b7
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265698"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="5614a-102">Créer un complément de volet de tâches Excel à l’aide de Vue</span><span class="sxs-lookup"><span data-stu-id="5614a-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="5614a-103">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de Vue et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="5614a-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5614a-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="5614a-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="5614a-105">Installez l’[interface de ligne de commande Vue](https://cli.vuejs.org/) globalement.</span><span class="sxs-lookup"><span data-stu-id="5614a-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="5614a-106">Génération d’une nouvelle application Vue</span><span class="sxs-lookup"><span data-stu-id="5614a-106">Generate a new Vue app</span></span>

<span data-ttu-id="5614a-p101">Utilisez l’interface de ligne de commande Vue pour générer une nouvelle application Vue. À partir du terminal, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="5614a-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="5614a-109">Ensuite, sélectionnez la présélection `default`.</span><span class="sxs-lookup"><span data-stu-id="5614a-109">Then select the `default` preset.</span></span> <span data-ttu-id="5614a-110">Si vous êtes invité à utiliser Yarn ou NPM comme package, vous pouvez choisir l’un ou l’autre.</span><span class="sxs-lookup"><span data-stu-id="5614a-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="5614a-111">Génération du fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="5614a-111">Generate the manifest file</span></span>

<span data-ttu-id="5614a-112">Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="5614a-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="5614a-113">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="5614a-113">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="5614a-114">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément en exécutant la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="5614a-114">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="5614a-115">Lorsque vous exécutez la commande `yo office`, il est possible que vous receviez des messages d’invite sur les règles de collecte de données de Yeoman et les outils CLI de complément Office.</span><span class="sxs-lookup"><span data-stu-id="5614a-115">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="5614a-116">Utilisez les informations fournies pour répondre aux invites comme vous l’entendez.</span><span class="sxs-lookup"><span data-stu-id="5614a-116">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="5614a-117">Si vous sélectionnez **Quitter** en réponse à la deuxième invite, vous devez réexécuter la commande `yo office` lorsque vous êtes prêt à créer votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="5614a-117">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="5614a-118">Lorsque vous y êtes invité, fournissez les informations suivantes pour créer votre projet de complément :</span><span class="sxs-lookup"><span data-stu-id="5614a-118">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="5614a-119">**Sélectionnez un type de projet :** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="5614a-119">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="5614a-120">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="5614a-120">**What do you want to name your add-in?**</span></span> `my-office-add-in`
    - <span data-ttu-id="5614a-121">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="5614a-121">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Générateur Yeoman](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="5614a-123">Une fois que vous avez terminé les étapes de l’Assistant, celui-ci crée un dossier `my-office-add-in` qui contient un fichier `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="5614a-123">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="5614a-124">Vous utiliserez le manifeste pour charger une version test et tester votre complément à la fin du Démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="5614a-124">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="5614a-125">Vous pouvez ignorer les *instructions suivantes* fournies par le générateur Yeoman une fois que le complément a été créé.</span><span class="sxs-lookup"><span data-stu-id="5614a-125">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="5614a-126">Les instructions détaillées de cet article fournissent tous les conseils nécessaires à l’exécution de ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="5614a-126">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="5614a-127">Sécurisation de l’application</span><span class="sxs-lookup"><span data-stu-id="5614a-127">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="5614a-128">Pour activer HTTPS pour votre application, créez un fichier `vue.config.js` dans le dossier racine du projet Vue avec le contenu suivant :</span><span class="sxs-lookup"><span data-stu-id="5614a-128">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="5614a-129">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="5614a-129">Update the app</span></span>

1. <span data-ttu-id="5614a-130">Ouvrez le fichier `public/index.html` et ajoutez la balise `<script>` suivante juste avant la balise `</head>` :</span><span class="sxs-lookup"><span data-stu-id="5614a-130">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="5614a-131">Ouvrez `src/main.js` et remplacez le contenu par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="5614a-131">Open `src/main.js` and replace the contents with the following code:</span></span>

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

3. <span data-ttu-id="5614a-132">Ouvrez `src/App.vue` et remplacez le contenu du fichier par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="5614a-132">Open `src/App.vue` and replace the file contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
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

## <a name="start-the-dev-server"></a><span data-ttu-id="5614a-133">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="5614a-133">Start the dev server</span></span>

1. <span data-ttu-id="5614a-134">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="5614a-134">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="5614a-135">Dans un navigateur web, accédez à `https://localhost:3000` (remarquez le `https`).</span><span class="sxs-lookup"><span data-stu-id="5614a-135">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="5614a-136">Si votre navigateur indique que le certificat de site n’est pas approuvé, vous devez [configurer votre ordinateur pour qu’il approuve le certificat](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="5614a-136">If your browser indicates that the site's certificate is not trusted, you will need to [configure your computer to trust the certificate](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span></span>

3. <span data-ttu-id="5614a-137">Lorsque la page sur `https://localhost:3000` est vide et qu’aucune erreur de certificat ne s’affiche, cela signifie qu’elle fonctionne.</span><span class="sxs-lookup"><span data-stu-id="5614a-137">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="5614a-138">L’application Vue est montée une fois qu’Office est initialisé, de sorte qu’elle affiche uniquement les éléments dans un environnement Excel.</span><span class="sxs-lookup"><span data-stu-id="5614a-138">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="5614a-139">Essayez</span><span class="sxs-lookup"><span data-stu-id="5614a-139">Try it out</span></span>

1. <span data-ttu-id="5614a-140">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5614a-140">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="5614a-141">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="5614a-141">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="5614a-142">Navigateur web : [Chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="5614a-142">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="5614a-143">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="5614a-143">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="5614a-144">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="5614a-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="5614a-146">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5614a-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5614a-147">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="5614a-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="5614a-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="5614a-149">Next steps</span></span>

<span data-ttu-id="5614a-150">Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide de Vue !</span><span class="sxs-lookup"><span data-stu-id="5614a-150">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="5614a-151">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="5614a-151">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5614a-152">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5614a-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="5614a-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5614a-153">See also</span></span>

* [<span data-ttu-id="5614a-154">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="5614a-154">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="5614a-155">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="5614a-155">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="5614a-156">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="5614a-156">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="5614a-157">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="5614a-157">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5614a-158">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="5614a-158">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="5614a-159">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="5614a-159">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
