---
title: Créer un complément de volet de tâches Excel à l’aide de Vue
description: ''
ms.date: 09/04/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9947852a586570345ba9f3dfe09340af6d01ace6
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715628"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="686ad-102">Créer un complément de volet de tâches Excel à l’aide de Vue</span><span class="sxs-lookup"><span data-stu-id="686ad-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="686ad-103">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de Vue et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="686ad-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="686ad-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="686ad-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="686ad-105">Installez l’[interface de ligne de commande Vue](https://cli.vuejs.org/) globalement.</span><span class="sxs-lookup"><span data-stu-id="686ad-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="686ad-106">Génération d’une nouvelle application Vue</span><span class="sxs-lookup"><span data-stu-id="686ad-106">Generate a new Vue app</span></span>

<span data-ttu-id="686ad-p101">Utilisez l’interface de ligne de commande Vue pour générer une nouvelle application Vue. À partir du terminal, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="686ad-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command and then answer the prompts as described below.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="686ad-109">Ensuite, sélectionnez la présélection `default`.</span><span class="sxs-lookup"><span data-stu-id="686ad-109">Then select the `default` preset.</span></span> <span data-ttu-id="686ad-110">Si vous êtes invité à utiliser Yarn ou NPM comme package, vous pouvez choisir l’un ou l’autre.</span><span class="sxs-lookup"><span data-stu-id="686ad-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="686ad-111">Génération du fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="686ad-111">Generate the manifest file</span></span>

<span data-ttu-id="686ad-112">Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="686ad-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="686ad-113">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="686ad-113">Navigate to your app folder.</span></span>

   ```command&nbsp;line
   cd my-add-in
   ```

2. <span data-ttu-id="686ad-p103">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="686ad-p103">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

   ```command&nbsp;line
   yo office
   ```

   ![Générateur Yeoman](../images/yo-office-manifest-only-vue.png)

   - <span data-ttu-id="686ad-117">**Sélectionnez un type de projet :** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="686ad-117">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
   - <span data-ttu-id="686ad-118">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="686ad-118">**What do you want to name your add-in?**</span></span> `my-office-add-in`
   - <span data-ttu-id="686ad-119">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="686ad-119">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="686ad-120">Une fois que vous avez terminé les étapes de l’Assistant, celui-ci crée un dossier `my-office-add-in` qui contient un fichier `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="686ad-120">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="686ad-121">Vous utiliserez le manifeste pour charger une version test et tester votre complément à la fin du Démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="686ad-121">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="686ad-122">Sécurisation de l’application</span><span class="sxs-lookup"><span data-stu-id="686ad-122">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="686ad-123">Pour activer HTTPS pour votre application, créez un fichier `vue.config.js` dans le dossier racine du projet Vue avec le contenu suivant :</span><span class="sxs-lookup"><span data-stu-id="686ad-123">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="686ad-124">Mettre à jour l’application</span><span class="sxs-lookup"><span data-stu-id="686ad-124">Update the app</span></span>

1. <span data-ttu-id="686ad-125">Ouvrez le fichier `public/index.html` et ajoutez la balise `<script>` suivante juste avant la balise `</head>` :</span><span class="sxs-lookup"><span data-stu-id="686ad-125">Open `public/index.html`, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="686ad-126">Ouvrez `src/main.js` et remplacez le contenu par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="686ad-126">Open the `src/main.js` file and replace it's contents with the following code.</span></span>

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

3. <span data-ttu-id="686ad-127">Ouvrez `src/App.vue` et remplacez le contenu du fichier par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="686ad-127">Open the `src/App.vue` file and replace it's contents with the following code.</span></span>

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

## <a name="start-the-dev-server"></a><span data-ttu-id="686ad-128">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="686ad-128">Start the dev server</span></span>

1. <span data-ttu-id="686ad-129">À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="686ad-129">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="686ad-130">Dans un navigateur web, accédez à `https://localhost:3000` (remarquez le `https`).</span><span class="sxs-lookup"><span data-stu-id="686ad-130">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="686ad-131">Si votre navigateur indique que le certificat de site n’est pas approuvé, vous devez [configurer votre ordinateur pour qu’il approuve le certificat](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="686ad-131">If your browser indicates that the site's certificate is not trusted, you will need to configure your computer to trust the certificate.</span></span>

3. <span data-ttu-id="686ad-132">Lorsque la page sur `https://localhost:3000` est vide et qu’aucune erreur de certificat ne s’affiche, cela signifie qu’elle fonctionne.</span><span class="sxs-lookup"><span data-stu-id="686ad-132">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="686ad-133">L’application Vue est montée une fois qu’Office est initialisé, de sorte qu’elle affiche uniquement les éléments dans un environnement Excel.</span><span class="sxs-lookup"><span data-stu-id="686ad-133">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="686ad-134">Essayez</span><span class="sxs-lookup"><span data-stu-id="686ad-134">Try it out</span></span>

1. <span data-ttu-id="686ad-135">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="686ad-135">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="686ad-136">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="686ad-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="686ad-137">Navigateur web : [Chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="686ad-137">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="686ad-138">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="686ad-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="686ad-139">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="686ad-139">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="686ad-141">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="686ad-141">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="686ad-142">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="686ad-142">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="686ad-144">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="686ad-144">Next steps</span></span>

<span data-ttu-id="686ad-145">Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide de Vue !</span><span class="sxs-lookup"><span data-stu-id="686ad-145">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="686ad-146">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="686ad-146">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="686ad-147">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="686ad-147">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="686ad-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="686ad-148">See also</span></span>

* [<span data-ttu-id="686ad-149">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="686ad-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="686ad-150">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="686ad-150">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="686ad-151">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="686ad-151">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="686ad-152">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="686ad-152">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
