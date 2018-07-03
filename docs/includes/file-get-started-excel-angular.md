# <a name="build-an-excel-add-in-using-angular"></a>Créer un complément Excel à l’aide d’Angular

Dans cet article, vous allez découvrir le processus de création d’un complément Excel à l’aide d’Angular et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

- Assurez-vous que vous avez la [configuration requise pour CLI Angular](https://github.com/angular/angular-cli#prerequisites) et installez tous les composants manquants.

- Installez [CLI Angular](https://github.com/angular/angular-cli) globalement. 

    ```bash
    npm install -g @angular/cli
    ```

- Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>Générer une nouvelle application Angular

Utilisez l’outil CLI Angular pour générer votre application Angular. À partir du terminal, exécutez la commande suivante :

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a>Génération du fichier manifeste

Le fichier manifeste d’un complément définit ses paramètres et ses fonctionnalités.

1. Accédez au dossier de votre application.

    ```bash
    cd my-addin
    ```

2. Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué ci-dessous.

    ```bash
    yo office 
    ```

    - **Choisissez un type de projet :** `Manifest`
    - **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ?** `Excel`


    Après avoir terminé l'assistant, un fichier manifeste et un fichier de ressources sont disponibles pour vous permettre de générer votre projet.

    ![Générateur Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Si vous êtes invité à remplacer **package.json**, répondez **Non** (ne pas remplacer).

## <a name="secure-the-app"></a>Sécurisation de l’application

[!include[HTTPS guidance](../includes/https-guidance.md)]

Pour ce démarrage rapide, vous pouvez utiliser les certificats fournis par le **générateur de compléments Office Yeoman**. Vous avez déjà installé le générateur globalement (comme demandé dans la section **Conditions préalables** de ce démarrage rapide). Vous n’avez donc qu’à copier les certificats situés dans l’emplacement d’installation global dans le dossier de votre application. La procédure suivante explique comment effectuer cette procédure.

1. À partir du terminal, exécutez la commande suivante pour identifier le dossier où les bibliothèques **npm** globales sont installées :

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > La première ligne de la sortie générée par cette commande spécifie le dossier où les bibliothèques **npm** globales sont installées.          
    
2. À l’aide de l’explorateur de fichiers, accédez au dossier `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`. À partir de cet emplacement, copiez le dossier `certs` dans votre presse-papiers.

3. Accédez au dossier racine de l’application Angular que vous avez créée à l’étape 1 de la section précédente et collez le dossier `certs` (qui se trouve dans votre presse-papiers) dans ce dossier.

## <a name="update-the-app"></a>Mettre à jour l’application

1. Dans votre éditeur de code, ouvrez **package.json** à la racine du projet. Modifiez le script `start` pour spécifier que le serveur doit s’exécuter à l’aide de SSL et du port 3000, puis enregistrez le fichier.

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. Ouvrez **.angular-cli.json** à la racine du projet. Modifiez l’objet **defaults** pour indiquer l’emplacement des fichiers de certificat et enregistrez le fichier.

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

3. Ouvrez **src/index.html**, ajoutez la balise `<script>` suivante immédiatement avant la balise `</head>`, puis enregistrez le fichier.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. Ouvrez **src/main.ts**, remplacez `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` par le code suivant, puis enregistrez le fichier. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. Ouvrez **src/polyfills.ts**, ajoutez la ligne de code suivante au-dessus de toutes les instructions `import` existantes, puis enregistrez le fichier.

    ```typescript
    import 'core-js/client/shim';
    ```

6. Dans **src/polyfills.ts**, supprimez les commentaires des lignes suivantes, puis enregistrez le fichier.

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

7. Ouvrez **src/app/app.component.html**, remplacez le contenu du fichier par le code HTML suivant, puis enregistrez le fichier. 

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

8. Ouvrez **src/app/app.component.css**, remplacez le contenu du fichier par le code CSS suivant, puis enregistrez le fichier.

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

9. Ouvrez **src/app/app.component.ts**, remplacez le contenu du fichier par le code suivant, puis enregistrez le fichier. 

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

## <a name="start-the-dev-server"></a>Démarrage du serveur de développement

1. À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.

    ```bash
    npm run start
    ```

2. Dans un navigateur web, accédez à `https://localhost:3000`. Si votre navigateur indique que le certificat du site n’est pas approuvé, vous devrez ajouter le certificat en tant que certificat approuvé. Consultez la rubrique relative à l’[ajout de certificats auto-signés en tant que certificats racine approuvés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour obtenir plus de détails.

    > [!NOTE]
    > Il est possible que le navigateur web Chrome continue d’indiquer que le certificat du site n’est pas approuvé, même si vous avez suivi les étapes décrites dans l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Vous pouvez ignorer ce message d’avertissement dans Chrome. Vérifiez tout de même que le certificat est approuvé en entrant `https://localhost:3000` dans Internet Explorer ou Microsoft Edge. 

3. Une fois que votre navigateur a chargé la page du complément sans erreurs de certificat, vous pouvez tester votre complément. 

## <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.

    - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

   
2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. Sélectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Excel à l’aide d’Angular ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts de base de l’API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Référence de l’API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
