---
title: Utiliser Vue pour créer un complément du volet de tâches Excel
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et de Vue pour Office.
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1686f9d9537718eb5ba56fa9ea7f0b4ccb7d65ec
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774440"
---
# <a name="use-vue-to-build-an-excel-task-pane-add-in"></a>Utiliser Vue pour créer un complément du volet de tâches Excel

Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de Vue et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Installez l’[interface de ligne de commande Vue](https://cli.vuejs.org/) globalement. À partir du terminal, exécutez la commande suivante.

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>Génération d’une nouvelle application Vue

Pour générer une nouvelle application Vue, utilisez l’interface de ligne de commande Vue.

```command&nbsp;line
vue create my-add-in
```

Ensuite, sélectionnez la `Default` prédéfinie pour « Vue 3 » (si vous préférez, choisissez « Vue 2 »).

## <a name="generate-the-manifest-file"></a>Génération du fichier manifeste

Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.

1. Accédez au dossier de votre application.

    ```command&nbsp;line
    cd my-add-in
    ```

1. Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément.

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > Lorsque vous exécutez la commande `yo office`, il est possible que vous receviez des messages d’invite sur les règles de collecte de données de Yeoman et les outils CLI de complément Office. Utilisez les informations fournies pour répondre aux invites en fonction des besoins. Si vous sélectionnez **Quitter** en réponse à la deuxième invite, vous devez réexécuter la commande `yo office` lorsque vous êtes prêt à créer votre projet de complément.

    Lorsque vous y êtes invité, fournissez les informations suivantes pour créer votre projet de complément.

    - **Sélectionnez un type de projet :** `Office Add-in project containing the manifest only`
    - **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

    ![Capture d’écran des invites d’interface de ligne de commande du générateur de compléments Yeoman Office pour les projets de fonctions personnalisées.](../images/yo-office-manifest-only-vue.png)

Une fois l’exécution terminée, l’Assistant crée un dossier **Mon complément Office** contenant un fichier **manifest.xml**. Vous utiliserez le manifeste pour charger une version test et tester votre complément à la fin du Démarrage rapide.

> [!TIP]
> Vous pouvez ignorer les *instructions suivantes* fournies par le générateur Yeoman une fois que le complément a été créé. Les instructions détaillées de cet article fournissent tous les conseils nécessaires à l’exécution de ce didacticiel.

## <a name="secure-the-app"></a>Sécurisation de l’application

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. Activez HTTPS pour votre application. Dans le dossier racine du projet Vue, créez un fichier **vue.config.js** avec le contenu suivant.

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

1. Installez les certificats du complément.

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé avec le générateur Yeoman contient un exemple de code pour un complément de volet Office de base. Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous. Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.

- Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément. Pour en savoir plus sur le fichier **manifest.xml**, consultez [manifeste XML des compléments Office](../develop/add-in-manifests.md).
- Le fichier **./src/App.vue** contient le balisage HTML du volet Office, le CSS appliqué au contenu du volet Office et le code de l’API JavaScript Office qui facilite l’interaction entre le volet Office et Excel.

## <a name="update-the-app"></a>Mettre à jour l’application

1. Ouvrez le fichier **./public/index.html** et ajoutez la balise `<script>` suivante juste avant la balise `</head>`.

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

1. Ouvrez **manifest.xml** et recherchez les balises `<bt:Urls>` dans la balise `<Resources>` . Recherchez la balise `<bt:Url>` avec l’ID `Taskpane.Url` et mettez à jour son attribut `DefaultValue`. La nouvelle `DefaultValue` est `https://localhost:3000/index.html`. La balise mise à jour entière doit correspondre à la ligne suivante.

   ```html
   <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
   ```

1. Ouvrez **./src/main.js** et remplacez le contenu par le code suivant.

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

1. Ouvrez **./src/App.vue** et remplacez le contenu du fichier par le code suivant.

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

## <a name="start-the-dev-server"></a>Démarrage du serveur de développement

1. Installer les dépendances.

     ```command&nbsp;line
    npm install
    ```

1. Démarrage du serveur de développement.

   ```command&nbsp;line
   npm run serve
   ```

1. Dans un navigateur web, accédez à `https://localhost:3000` (remarquez le `https`). Si la page sur `https://localhost:3000` est vide et qu’aucune erreur de certificat ne s’affiche, cela signifie qu’elle fonctionne. L’application Vue est montée une fois qu’Office est initialisé, de sorte qu’elle affiche uniquement les éléments dans un environnement Excel.

## <a name="try-it-out"></a>Essayez

1. Exécutez votre complément et chargez-le de manière indépendante dans Excel. Suivez les instructions pour la plateforme que vous utiliserez :

   - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Navigateur web : [Chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

1. Ouvrez le volet Office du complément dans Excel. Dans l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches**.

   ![Capture d’écran du menu Accueil d’ Excel, avec le bouton Afficher le volet Office mis en évidence.](../images/excel-quickstart-addin-2a.png)

1. Sélectionnez une plage de cellules dans la feuille de calcul.

1. Définissez la couleur de la plage sélectionnée sur vert. Dans le volet Office de votre complément, choisissez le bouton **Définir la couleur** .

   ![Capture d’écran d’ Excel avec le volet Office Complément ouvert.](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément du volet Office Excel à l’aide de Vue ! Maintenant, apprenez-en davantage sur les fonctionnalités d’un complément Excel et créez un complément plus complexe en suivant le didacticiel sur les compléments Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](../excel/excel-add-ins-core-concepts.md)
- [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
