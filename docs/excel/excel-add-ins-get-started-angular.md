# <a name="build-an-excel-add-in-using-angular"></a>Créer un complément Excel à l’aide d’Angular

Dans cet article, vous allez découvrir le processus de création d’un complément Excel à l’aide d’Angular et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

Si vous ne l’avez pas fait précédemment, installez les outils suivants :

1. Assurez-vous que vous avez la [configuration requise pour CLI Angular](https://github.com/angular/angular-cli#prerequisites) et installez tous les composants manquants.

2. Installez [CLI Angular](https://github.com/angular/angular-cli) globalement. 

    ```bash
    npm install -g @angular/cli
    ```

3. Installez [Yeoman](https://github.com/yeoman/yo) et [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>Générer une nouvelle application Angular

Utilisez l’outil CLI Angular pour générer votre application Angular. À partir du terminal, exécutez la commande suivante :

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>Générer le fichier manifeste et charger une version test du complément

Le fichier manifeste d’un complément définit ses paramètres et ses fonctionnalités.

1. Accédez au dossier de votre application.

    ```bash
    cd my-addin
    ```

2. Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué dans la capture d’écran ci-dessous.

    ```bash
    yo office
    ```
    ![Générateur Yeoman](../../images/yo-office.png)
    > **Remarque** : si vous êtes invité à remplacer **package.json**, répondez **Non** (ne pas remplacer).

3. Ouvrez le fichier manifeste (à savoir, le fichier dans le répertoire racine de votre application dont le nom se termine par « manifest.xml »). Remplacez toutes les occurrences de `https://localhost:3000` par `http://localhost:4200`, puis enregistrez le fichier.

    > **Remarque** : assurez-vous que vous avez défini le protocole sur **http** en plus de la définition du numéro de port sur **4200**.

4. Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.

    - Windows : [Chargement de compléments Office pour des tests sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>Mettre à jour l’application

1. Ouvrez **src/index.html**, ajoutez la balise `<script>` suivante immédiatement avant la balise `</head>`, puis enregistrez le fichier.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Ouvrez **src/main.ts**, remplacez `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` par le code suivant, puis enregistrez le fichier. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

3. Ouvrez **src/polyfills.ts**, ajoutez la ligne de code suivante au-dessus de toutes les instructions `import` existantes, puis enregistrez le fichier.

    ```typescript
    import 'core-js/client/shim';
    ```

4. Dans **src/polyfills.ts**, supprimez les commentaires des lignes suivantes, puis enregistrez le fichier.

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

5. Ouvrez **src/app/app.component.html**, remplacez le contenu du fichier par le code HTML suivant, puis enregistrez le fichier. 

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
            <button (click)="onColorMe()">Color Me</button>
        </div>
    </div>
    ```

6. Ouvrez **src/app/app.component.css**, remplacez le contenu du fichier par le code CSS suivant, puis enregistrez le fichier.

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

7. Ouvrez **src/app/app.component.ts**, remplacez le contenu du fichier par le code suivant, puis enregistrez le fichier. 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onColorMe() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="try-it-out"></a>Essayez !

1. À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.

    ```bash
    npm start
    ```

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Bouton Complément Excel](../../images/excel_quickstart_addin_2a.png)

3. Choisissez le bouton **Indiquer une couleur** dans le volet Office pour définir la couleur de la plage sélectionnée sur vert.

    ![Complément Excel](../../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Excel à l’aide d’Angular ! Apprenez-en davantage sur les [concepts fondamentaux](excel-add-ins-core-concepts.md) de la création de compléments Excel.

## <a name="additional-resources"></a>Ressources supplémentaires

* [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Référence de l’API JavaScript pour Excel](../../reference/excel/excel-add-ins-reference-overview.md)
