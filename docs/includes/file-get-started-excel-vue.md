# <a name="build-an-excel-add-in-using-vue"></a>Développement d’un complément Excel à l’aide de Vue

Cet article décrit le processus de création d’un complément Excel à l’aide de Vue et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org)

- Installez l’[interface de ligne de commande Vue](https://github.com/vuejs/vue-cli) globalement.

    ```bash
    npm install -g vue-cli
    ```

- Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-vue-app"></a>Génération d’une nouvelle application Vue

Utilisez l’interface de ligne de commande Vue pour générer une nouvelle application Vue. À partir du terminal, exécutez la commande suivante, puis répondez aux invites comme décrit ci-dessous.

```bash
vue init webpack my-add-in
```

Lorsque vous répondez aux invites générées par la commande précédente, remplacez les réponses par défaut des 3 invites ci-dessous. Vous pouvez accepter les réponses par défaut de toutes les autres invites.

- **Installer vue-router ?** `No`
- **Configurer des tests d’unités :** `No`
- **Configurer des tests e2e avec Nightwatch ?** `No`

![Invites de l’interface de ligne de commande Vue](../images/vue-cli-prompts.png)

## <a name="generate-the-manifest-file"></a>Génération du fichier manifeste

Chaque complément nécessite un fichier manifeste pour définir ses paramètres et ses fonctionnalités.

1. Accédez au dossier de votre application.

    ```bash
    cd my-add-in
    ```

2. Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué ci-dessous.

    ```bash
    yo office 
    ```

    - **Sélectionnez un type de projet :** `Office Add-in containing the manifest only`
    - **Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :**`Excel`

    ![Générateur Yeoman](../images/yo-office.png)
    
    Une fois l’Assistant exécuté, le générateur crée le fichier manifeste.

## <a name="secure-the-app"></a>Sécurisation de l’application

[!include[HTTPS guidance](../includes/https-guidance.md)]

Pour activer HTTPS pour votre application, ouvrez le fichier **package.json** dans le dossier racine du projet Vue, modifiez le script `dev` pour ajouter le marqueur `--https` et enregistrez le fichier.

```json
"dev": "webpack-dev-server --https --inline --progress --config build/webpack.dev.conf.js"
```

## <a name="update-the-app"></a>Mettre à jour l’application

1. Dans votre éditeur de code, ouvrez le dossier **My Office Add-in** créé par Yo Office à la racine de votre projet Vue. Dans ce dossier, vous verrez le fichier manifeste qui définit les paramètres de votre complément : **manifest.xml**.

2. Ouvrir le fichier manifeste, remplacez toutes les occurrences de `https://localhost:3000` par `https://localhost:8080` et enregistrez le fichier.

3. Ouvrez le fichier **index.html** (situé à la racine de votre projet Vue), ajoutez la balise `<script>` suivante immédiatement avant la balise `</head>`, puis enregistrez le fichier.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

3. Ouvrez **src/main.js** et *supprimez* le bloc de code suivant :

    ```js
    new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
    })
    ```
    
    Ajoutez le code suivant à ce même emplacement, puis enregistrez le fichier. 
                                                         
    ```js
    const Office = window.Office
    Office.initialize = () => {
      new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
      })
    }
    ```

4. Ouvrez **src/App.vue**, remplacez le contenu du fichier par le code suivant, ajoutez un saut de ligne à la fin du fichier (c’est-à-dire, après la balise `</style>`) et enregistrez le fichier. 

    ```html
    <template>
    <div id="app">
        <div id="content">
        <div id="content-header">
            <div class="padding">
            <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br/>
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
        onSetColor () {
          window.Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange()
            range.format.fill.color = 'green'
            await context.sync()
          })
        }
      }
    }
    </script>

    <style>
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
    </style>
    ```

## <a name="start-the-dev-server"></a>Démarrage du serveur de développement

1. À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.

    ```bash
    npm start
    ```

2. Dans un navigateur web, accédez à `https://localhost:8080`. Si votre navigateur indique que le certificat de site n’est pas approuvé, vous devez configurer votre ordinateur pour qu’il approuve le certificat. 

3. Une fois que votre navigateur a chargé la page du complément sans erreurs de certificat, vous pouvez tester votre complément. 

## <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.

    - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. Sélectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Excel à l’aide de Vue ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
