---
title: 'Didacticiel : créer un complément de composition de message Outlook'
description: Dans ce didacticiel, vous allez créer un complément Outlook qui insère des informations GitHub dans le corps d'un nouveau message.
ms.date: 02/01/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 56def561fee6525c6daa73fe1153f220bae503c3
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238098"
---
# <a name="tutorial-build-a-message-compose-outlook-add-in"></a>Didacticiel : créer un complément de composition de message Outlook

Ce didacticiel vous apprend à créer un complément Outlook qui peut être utilisé pour dans le mode composer un message pour insérer du contenu dans le corps d’un message.

Dans ce didacticiel, vous allez :

> [!div class="checklist"]
>
> - Créer un projet de complément Outlook
> - Définir des boutons qui s’afficheront dans la fenêtre composer un message
> - Implémenter une expérience de première exécution qui collecte des informations de l’utilisateur et extrait les données à partir d’un service externe
> - Implémenter un bouton de l’interface utilisateur qui appelle une fonction
> - Implémenter un volet des tâches qui insère du contenu dans le corps d’un message

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))

- La dernière version de[Yeoman](https://github.com/yeoman/yo) et du [Générateur Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante via l’invite de commande.

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Même si vous avez précédemment installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version de npm.

- Outlook 2016 ou plus récent sur Windows (connecté à un compte Microsoft 365) ou Outlook sur le web

- Un compte[GitHub](https://www.github.com) 

## <a name="setup"></a>Configuration

Le complément que vous allez créer dans ce didacticiel lit les[gists](https://gist.github.com) à partir du compte utilisateur GitHub et ajoute le gist sélectionné dans le corps d’un message. Procédez comme suit pour créer deux nouveaux gists que vous pouvez utiliser pour tester le complément que vous allez créer.

1. [Connectez-vous à GitHub](https://github.com/login).

1. [Créer une nouveau gist](https://gist.github.com).

    - Dans la zone **Description gist...**, entrez **Hello World Markdown**.

    - Dans la zone **Nom de fichier incluant l’extension...**, entrez **test.md**.

    - Ajoutez la démarque suivante à la zone de texte multiligne.

        ```markdown
        # Hello World

        This is content converted from Markdown!

        Here's a JSON sample:

          ```json
          {
            "foo": "bar"
          }
          ```
        ```

    - Sélectionnez le bouton **créer un gist public**.

1. [Créer un nouveau gist](https://gist.github.com).

    - Dans la zone **Description gist...**, entrez **Hello World Html**.

    - Dans la zone **Nom de fichier incluant l’extension...**, entrez **test.html**.

    - Ajoutez la démarque suivante à la zone de texte multiligne.

        ```HTML
        <html>
          <head>
            <style>
            h1 {
              font-family: Calibri;
            }
            </style>
          </head>
          <body>
            <h1>Hello World!</h1>
            <p>This is a test</p>
          </body>
        </html>
        ```

    - Sélectionnez le bouton **créer un gist public**.

## <a name="create-an-outlook-add-in-project"></a>Créer un projet de complément Outlook

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Sélectionnez un type de projet** - `Office Add-in Task Pane project`

    - **Sélectionnez un type de script** - `JavaScript`

    - **Comment souhaitez-vous nommer votre complément ?** - `Git the gist`

    - **Quelle application client Office voulez-vous prendre en charge ?** - `Outlook`

    ![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande](../images/yeoman-prompts-2.png)

    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Accédez au registre racine du projet.

    ```command&nbsp;line
    cd "Git the gist"
    ```

1. Ce complément utilise les bibliothèques suivantes.

    - Bibliothèque [Showdown](https://github.com/showdownjs/showdown) pour convertir Markdown en HTML.
    - Bibliothèque [URI.js](https://github.com/medialize/URI.js) pour créer des URL relatives.
    - Bibliothèque [jQuery](https://jquery.com/) pour simplifier les interactions DOM.

     Pour installer ces outils pour votre projet, exécutez la commande suivante dans le répertoire racine du projet.

    ```command&nbsp;line
    npm install showdown urijs jquery --save
    ```

### <a name="update-the-manifest"></a>Mise à jour du manifeste

Le manifeste d’un complément contrôle la manière dont il apparaît dans Outlook. Il définit la façon dont le complément est affiché dans la liste des compléments, les boutons qui apparaissent sur le ruban, et il configure les URL pour les fichiers HTML et JavaScript utilisés par le complément.

#### <a name="specify-basic-information"></a>Spécifiez les informations de base

Effectuez les mises à jour suivantes dans le fichier **manifest.xml** pour spécifier les informations de base du complément.

1. Recherchez l’élément `ProviderName`et remplacez la valeur par défaut par le nom de votre société.

    ```xml
    <ProviderName>Contoso</ProviderName>
    ```

1. Recherchez l’`Description` élément, remplacez la valeur par défaut avec une description du complément et enregistrez le fichier.

    ```xml
    <Description DefaultValue="Allows users to access their GitHub gists."/>
    ```

#### <a name="test-the-generated-add-in"></a>Tester le complément généré

Avant d’aller plus loin, nous allons tester le complément base créé par le générateur pour confirmer que le projet est correctement configuré.

> [!NOTE]
> Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman. Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution) et votre complément est chargé.

    ```command&nbsp;line
    npm start
    ```

1. Dans Outlook, ouvrez un message existant et sélectionnez le bouton **Afficher le volet Office**. Si tout est configuré correctement, le volet des tâches va s’ouvrir et afficher la page d’accueil du complément.

    ![Capture d’écran du bouton « Afficher le volet Office » et de la git volet Office ajouté par l’échantillon](../images/button-and-pane.png)

## <a name="define-buttons"></a>Définir des boutons

À présent que vous avez vérifié que le complément base fonctionne, vous pouvez le personnaliser pour ajouter davantage de fonctionnalités. Par défaut, le manifeste définit uniquement les boutons de la fenêtre de lecture de message. Nous allons mettre à jour le manifeste pour supprimer les boutons de la fenêtre de lecture de message et définir deux nouveaux boutons pour la fenêtre composer un message :

- **Insérer un gist**: bouton qui ouvre un le volet des tâches

- **Insérer gist par défaut**: bouton qui appelle une fonction

### <a name="remove-the-messagereadcommandsurface-extension-point"></a>Supprimer le point d’extension MessageReadCommandSurface

Ouvrir le fichier **manifest.xml** et rechercher l’`ExtensionPoint` élément avec un type `MessageReadCommandSurface`. Supprimer cet `ExtensionPoint` élément (y compris sa balise de fermeture) pour supprimer les boutons de la fenêtre de lecture de message.

### <a name="add-the-messagecomposecommandsurface-extension-point"></a>Supprimer le point d’extension MessageComposeCommandSurface

Recherchez la ligne dans le manifeste qui lit `</DesktopFormFactor>`. Situé immédiatement avant cette ligne, insérez le balisage XML suivant. Notez les points suivants concernant ce balisage.

- L’élément `ExtensionPoint` avec `xsi:type="MessageComposeCommandSurface"` indique que vous définissez des boutons à ajouter à la fenêtre de composition d’un message.

- En utilisant un élément `OfficeTab` avec `id="TabDefault"`, vous indiquez que vous voulez ajouter des boutons à l’onglet par défaut dans le ruban.

- L’élément `Group` définit le regroupement de nouveaux boutons, avec une étiquette définie par la ressource `groupLabel`.

- Le premier élément `Control` contient un élément `Action` avec `xsi:type="ShowTaskPane"`, afin que le bouton ouvre un volet des tâches.

- Le deuxième élément `Control` contient un élément `Action` avec `xsi:type="ExecuteFunction"`, afin que le bouton appelle une fonction JavaScript contenue dans le fichier de fonction.

```xml
<!-- Message Compose -->
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgComposeCmdGroup">
      <Label resid="GroupLabel"/>
      <Control xsi:type="Button" id="msgComposeInsertGist">
        <Label resid="TaskpaneButton.Label"/>
        <Supertip>
          <Title resid="TaskpaneButton.Title"/>
          <Description resid="TaskpaneButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="Taskpane.Url"/>
        </Action>
      </Control>
      <Control xsi:type="Button" id="msgComposeInsertDefaultGist">
        <Label resid="FunctionButton.Label"/>
        <Supertip>
          <Title resid="FunctionButton.Title"/>
          <Description resid="FunctionButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
          <FunctionName>insertDefaultGist</FunctionName>
        </Action>
      </Control>
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

### <a name="update-resources-in-the-manifest"></a>Ressources de mise à jour dans le fichier manifeste

Le code précédent fait référence à des étiquettes, des info-bulles et des URL que vous devez définir avant que le manifeste ne soit valide. Vous devez spécifier ces informations dans la section `Resources` du manifeste.

1. Recherchez l’élément `Resources` dans le fichier manifeste, puis supprimez entièrement l’élément (balise de fermeture comprise).

1. À ce même emplacement, ajoutez le balisage suivant pour remplacer l’élément `Resources` que vous venez de supprimer.

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Git the gist"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Insert gist"/>
        <bt:String id="TaskpaneButton.Title" DefaultValue="Insert gist"/>
        <bt:String id="FunctionButton.Label" DefaultValue="Insert default gist"/>
        <bt:String id="FunctionButton.Title" DefaultValue="Insert default gist"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Displays a list of your gists and allows you to insert their contents into the current message."/>
        <bt:String id="FunctionButton.Tooltip" DefaultValue="Inserts the content of the gist you mark as default into the current message."/>
      </bt:LongStrings>
    </Resources>
    ```

1. Enregistrez les modifications dans le manifeste.

### <a name="reinstall-the-add-in"></a>Réinstallez le complément.

Étant donné que vous avez installé le complément à partir d’un fichier, vous devez le réinstaller afin que les modifications soient prises en compte.

1. Suivez les instructions pour supprimer **Git the gist** des [compléments sideloaded](../outlook/sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in).

1. Fermer la fenêtre **Mes compléments**.

1. Le bouton personnalisé doit disparaître du ruban temporairement.

1. Suivez les instructions de [Charger compléments Outlook pour les tests](../outlook/sideload-outlook-add-ins-for-testing.md) pour réinstaller le complément à l’aide du fichier mis à jour **manifest.xml**.

Une fois le complément réinstallé, vous pouvez vérifier qu’il a été correctement installé en consultant les commandes **Insérer gist** et **Insérer gist par défaut** dans le fenêtre de composition du message. Notez que rien ne se produit si vous sélectionnez un des ces éléments, car vous n’avez pas encore terminé de générer ce complément.

- Si vous exécutez ce complément dans Outlook 2016 ou versions ultérieures sur Windows, vous devriez voir deux nouveaux boutons dans le ruban de la fenêtre de composition d’un message : **Insérer gist** et **Insérer gist par défaut**.

    ![Capture d’écran du menu de dépassement de ruban dans Outlook sur Windows avec les boutons du complément mis en évidence](../images/add-in-buttons-in-windows.png)

- Si vous exécutez ce complément dans Outlook sur le web, vous devriez voir apparaître un nouveau bouton en bas de la fenêtre de composition d’un message. Sélectionnez ce bouton pour afficher les options **Insérer gist** et **Insérer gist par défaut**.

    ![Capture d’écran du formulaire composer message dans Outlook sur le web avec le bouton complément et menu contextuel mis en évidence](../images/add-in-buttons-in-owa.png)

## <a name="implement-a-first-run-experience"></a>Mettre en œuvre une expérience de première exécution

Ce complément doit être en mesure de lire les gists du compte d’utilisateur GitHub et d’identifier lequel l’utilisateur a choisi en tant que gist par défaut. Pour atteindre ces objectifs, le complément doit inviter l’utilisateur à fournir son nom d’utilisateur GitHub et choisir un gist par défaut parmi leur collection de gists existants. Suivez les étapes décrites dans cette section pour implémenter une expérience de première exécution qui affiche une boîte de dialogue pour collecter ces informations à partir de l’utilisateur.

### <a name="collect-data-from-the-user"></a>Collecter les données d’un utilisateur

Commençons par créer l’interface utilisateur pour la boîte de dialogue. Dans le dossier **./src**, créez un sous-dossier nommé **settings**. Dans le dossier **./src/settings**, créez un fichier nommé **dialog.html** et ajoutez le balisage suivant pour définir un formulaire très simple avec une entrée de texte pour un nom d’utilisateur GitHub et une liste vide pour gists qui sera renseignée via JavaScript.

```html
<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <title>Settings</title>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

  <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

  <!-- Template styles -->
  <link href="dialog.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l">
  <main>
    <section class="ms-font-m ms-fontColor-neutralPrimary">
      <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Info"></i>
          </div>
          <div class="ms-MessageBar-text">
            Oops! It looks like you haven't configured <strong>Git the gist</strong> yet.
            <br/>
            Please configure your GitHub username and select a default gist, then try that action again!
          </div>
        </div>
      </div>
      <div class="ms-font-xxl">Settings</div>
      <div class="ms-Grid">
        <div class="ms-Grid-row">
          <div class="ms-TextField">
            <label class="ms-Label">GitHub Username</label>
            <input class="ms-TextField-field" id="github-user" type="text" value="" placeholder="Please enter your GitHub username">
          </div>
        </div>
        <div class="error-display ms-Grid-row">
          <div class="ms-font-l ms-fontWeight-semibold">An error occurred:</div>
          <pre><code id="error-text"></code></pre>
        </div>
        <div class="gist-list-container ms-Grid-row">
          <div class="list-title ms-font-xl ms-fontWeight-regular">Choose Default Gist</div>
          <form>
            <div id="gist-list">
            </div>
          </form>
        </div>
      </div>
      <div class="ms-Dialog-actions">
        <div class="ms-Dialog-actionsRight">
          <button class="ms-Dialog-action ms-Button ms-Button--primary" id="settings-done" disabled>
            <span class="ms-Button-label">Done</span>
          </button>
        </div>
      </div>
    </section>
  </main>
  <script type="text/javascript" src="../../node_modules/core-js/client/core.js"></script>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="dialog.js"></script>
</body>

</html>
```

Ensuite, créez un fichier dans le dossier **./src/settings** nommé **dialog.css** et ajoutez le code suivant pour spécifier les styles utilisés par **dialog.html**.

```CSS
section {
  margin: 10px 20px;
}

.not-configured-warning {
  display: none;
}

.error-display {
  display: none;
}

.gist-list-container {
  margin: 10px -8px;
  display: none;
}

.list-title {
  border-bottom: 1px solid #a6a6a6;
  padding-bottom: 5px;
}

ul {
  margin-top: 10px;
}

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}
```

Maintenant que vous avez défini la boîte de dialogue interface utilisateur, vous pouvez écrire du code pour l’utiliser. Créez un fichier dans le dossier **./src/settings** nommé **dialog.js** et ajoutez le code suivant. Notez que ce code utilise jQuery pour enregistrer des événements et la fonction `messageParent` pour renvoyer les choix de l’utilisateur à l’appelant.

```js
(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      if (window.location.search) {
        // Check if warning should be displayed.
        var warn = getParameterByName('warn');
        if (warn) {
          $('.not-configured-warning').show();
        } else {
          // See if the config values were passed.
          // If so, pre-populate the values.
          var user = getParameterByName('gitHubUserName');
          var gistId = getParameterByName('defaultGistId');

          $('#github-user').val(user);
          loadGists(user, function(success){
            if (success) {
              $('.ms-ListItem').removeClass('is-selected');
              $('input').filter(function() {
                return this.value === gistId;
              }).addClass('is-selected').attr('checked', 'checked');
              $('#settings-done').removeAttr('disabled');
            }
          });
        }
      }

      // When the GitHub username changes,
      // try to load gists.
      $('#github-user').on('change', function(){
        $('#gist-list').empty();
        var ghUser = $('#github-user').val();
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', function() {
        var settings = {};

        settings.gitHubUserName = $('#github-user').val();

        var selectedGist = $('.ms-ListItem.is-selected');
        if (selectedGist) {
          settings.defaultGistId = selectedGist.val();

          sendMessage(JSON.stringify(settings));
        }
      });
    });
  };

  // Load gists for the user using the GitHub API
  // and build the list.
  function loadGists(user, callback) {
    getUserGists(user, function(gists, error){
      if (error) {
        $('.gist-list-container').hide();
        $('#error-text').text(JSON.stringify(error, null, 2));
        $('.error-display').show();
        if (callback) callback(false);
      } else {
        $('.error-display').hide();
        buildGistList($('#gist-list'), gists, onGistSelected);
        $('.gist-list-container').show();
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
    $('.not-configured-warning').hide();
    $('#settings-done').removeAttr('disabled');
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
```

#### <a name="update-webpack-config-settings"></a>Mettre à jour les paramètres de configuration webapck

Enfin, ouvrez le fichier **webpack.config.js** situé dans le répertoire racine du projet et procédez comme suit.

1. Recherchez l’objet `entry` dans l’objet `config` et ajoutez une nouvelle entrée pour `dialog`.

    ```js
    dialog: "./src/settings/dialog.js"
    ```

    Lorsque c’est chose faite, le nouvel objet `entry` se présente comme suit :

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      dialog: "./src/settings/dialog.js"
    },
    ```

1. Recherchez la matrice `plugins` au sein de l’objet `config`. Dans la matrice `patterns` de l’objet `new CopyWebpackPlugin` , ajoutez une nouvelle entrée après l’entrée de `taskpane.css` .

    ```js
    {
      to: "dialog.css",
      from: "./src/settings/dialog.css"
    },
    ```

    Lorsque c’est chose faite, l’objet `new CopyWebpackPlugin` se présente comme suit :

    ```js
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        },
        {
          to: "dialog.css",
          from: "./src/settings/dialog.css"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
    ```

1. Recherchez la matrice `plugins` dans l’objet `config` et ajoutez ce nouvel objet à la fin de cette matrice.

    ```js
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: "./src/settings/dialog.html",
      chunks: ["polyfill", "dialog"]
    })
    ```

    Lorsque c’est chose faite, la nouvelle matrice `plugins` se présente comme suit :

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        },
        {
          to: "dialog.css",
          from: "./src/settings/dialog.css"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/settings/dialog.html",
        chunks: ["polyfill", "dialog"]
      })
    ],
    ```

1. Si le serveur web est en cours d’exécution, fermez la fenêtre de commande de nœud.

1. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

1. Exécutez la commande suivante pour démarrer le serveur web et ajouter votre module.

    ```command&nbsp;line
    npm start
    ```

### <a name="fetch-data-from-github"></a>Récupérer des données à partir de GitHub

Le fichier **dialog.js** que vous venez de créer spécifie que le complément doit charger les gists lorsque l’`change` événement se déclenche pour le champ nom d’utilisateur GitHub. Pour récupérer les gists de l’utilisateur à partir de GitHub, vous utiliserez le [API GitHub Gists](https://developer.github.com/v3/gists/).

Dans le dossier **./src**, créez un nouveau sous-dossier nommé **helpers**. Dans le dossier **./src/helpers**, créez un fichier nommé **gist-api.js** et ajoutez le code suivant pour récupérer les gists de l’utilisateur à partir de GitHub et créer la liste des gists.

```js
function getUserGists(user, callback) {
  var requestUrl = 'https://api.github.com/users/' + user + '/gists';

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gists){
    callback(gists);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function(gist) {

    var listItem = $('<div/>')
      .appendTo(parent);

    var radioItem = $('<input>')
      .addClass('ms-ListItem')
      .addClass('is-selectable')
      .attr('type', 'radio')
      .attr('name', 'gists')
      .attr('tabindex', 0)
      .val(gist.id)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-primaryText')
      .text(gist.description)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-secondaryText')
      .text(' - ' + buildFileList(gist.files))
      .appendTo(listItem);

    var updated = new Date(gist.updated_at);

    var desc = $('<span/>')
      .addClass('ms-ListItem-tertiaryText')
      .text(' - Last updated ' + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on('click', clickFunc);
  });  
}

function buildFileList(files) {

  var fileList = '';

  for (var file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ', ';
      }

      fileList = fileList + files[file].filename + ' (' + files[file].language + ')';
    }
  }

  return fileList;
}
```

> [!NOTE]
> Vous avez sans doute remarqué qu’il n’existe pas de bouton pour appeler la boîte de dialogue Paramètres. Au lieu de cela, le complément vérifie si cela a été configuré lorsque l’utilisateur sélectionne le bouton **Insérer gist par défaut** ou le bouton **Insérer gist**. Si le complément n'a pas encore été configuré, la boîte de dialogue Paramètres invite l’utilisateur à configurer avant de continuer.

## <a name="implement-a-ui-less-button"></a>Implémentation d’un bouton sans interface utilisateur

Le bouton **Insérer gist par défaut** de ce complément est un bouton sans interface utilisateur qui appelera une fonction JavaScript, plutôt que d’ouvrir un volet des tâches comme de nombreux boutons de complément le font. Lorsque l’utilisateur sélectionne le bouton **Insérer gist par défaut**, la fonction JavaScript correspondante vérifie si le complément a été configuré.

- Si le complément a déjà été configuré, la fonction chargera le contenu du gist que l’utilisateur a sélectionné par défaut et l’insérera dans le corps du message.

- Si le complément n'a pas encore été configuré, la boîte de dialogue Paramètres invitera l’utilisateur à fournir les informations nécessaires. 

### <a name="update-the-function-file-html"></a>Mettre à jour le fichier de fonction (HTML)

Une fonction appelée par un bouton sans interface utilisateur doit être définie dans le fichier de fonction spécifié par l’élément `FunctionFile` dans le manifeste pour le facteur de formulaire correspondant. Le manifeste de ce complément spécifie `https://localhost:3000/commands.html` comme fichier de fonction.

Ouvrez le fichier **./src/commands/commands.html** et remplacez tout le contenu par le balisage suivant.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <script type="text/javascript" src="../node_modules/jquery/dist/jquery.js"></script>
    <script type="text/javascript" src="../node_modules/showdown/dist/showdown.min.js"></script>
    <script type="text/javascript" src="../node_modules/urijs/src/URI.min.js"></script>
    <script type="text/javascript" src="../src/helpers/addin-config.js"></script>
    <script type="text/javascript" src="../src/helpers/gist-api.js"></script>
</head>

<body>
  <!-- NOTE: The body is empty on purpose. Since functions in commands.js are
       invoked via a button, there is no UI to render. -->
</body>

</html>
```

### <a name="update-the-function-file-javascript"></a>Mettre à jour le fichier de fonction (JavaScript)

Ouvrez le fichier **./src/commands/commands.js** et remplacez tout le contenu par le code suivant. Notez que si la `insertDefaultGist` fonction détermine que le complément n'a pas encore été configuré, elle ajoute le `?warn=1` paramètre à l’URL de la boîte de dialogue. Cette opération permet à la boîte de dialogue Paramètres de restituer la barre des messages définie dans **./settings/dialog.html**, pour transmettre à l’utilisateur pourquoi il voit la boîte de dialogue.

```js
var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function (reason) {
};

// Add any UI-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
    type: 'errorMessage',
    message: error
  }, function(result){
  });
}

var settingsDialog;

function insertDefaultGist(event) {

  config = getConfig();

  // Check if the add-in has been configured.
  if (config && config.defaultGistId) {
    // Get the default gist content and insert.
    try {
      getGist(config.defaultGistId, function(gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(content,
                {coercionType: Office.CoercionType.Html}, function(result) {
                  event.completed();
              });
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }

  } else {
    // Save the event object so we can finish up later.
    btnEvent = event;
    // Not configured yet, display settings dialog with
    // warn=1 to display warning.
    var url = new URI('../src/settings/dialog.html?warn=1').absoluteTo(window.location).toString();
    var dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    });
  }
}

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function(result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}

function getGlobal() {
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window :
    (typeof global !== "undefined") ? global :
    undefined;
}

var g = getGlobal();

// The add-in command functions need to be available in global scope.
g.insertDefaultGist = insertDefaultGist;
```

### <a name="create-a-file-to-manage-configuration-settings"></a>Créer un fichier pour gérer les paramètres de configuration

Le fichier fonction HTML fait référence à un fichier nommé **addin-config.js**, qui n’existe pas encore. Créez un fichier nommé **addin-config.js** dans le dossier **./src/helpers** et ajoutez le code suivant. Ce code utilise l’[objet RoamingSettings](/javascript/api/outlook/office.RoamingSettings) pour obtenir et définir les valeurs de configuration.

```js
function getConfig() {
  var config = {};

  config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
  config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
  Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);

  Office.context.roamingSettings.saveAsync(callback);
}
```

### <a name="create-new-functions-to-process-gists"></a>Créer de nouvelles fonctions pour traiter les gists

Ensuite, ouvrez le fichier **./src/helpers/gist-api.js** et ajoutez les fonctions suivantes. Veuillez prendre en compte les éléments suivants:

- Si le gist contient du HTML, le complément insère le code HTML tel quel dans le corps du message.

- Si le gist contient Markdown, le complément utilisera la bibliothèque[Showdown](https://github.com/showdownjs/showdown) pour convertir le Markdown en HTML, puis insérera le code HTML qui en résulte dans le corps du message.

- Si le gist contient autre chose que du HTML ou Markdown, le complément l’insère dans le corps du message comme un extrait de code.

```js
function getGist(gistId, callback) {
  var requestUrl = 'https://api.github.com/gists/' + gistId;

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gist){
    callback(gist);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (var filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      var file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case 'HTML':
            // Insert as-is.
            callback(file.content);
            break;
          case 'Markdown':
            // Convert Markdown to HTML.
            var converter = new showdown.Converter();
            var html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            var codeBlock = '<pre><code>';
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + '</code></pre>';
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, 'No suitable file found in the gist');
}
```

### <a name="test-the-button"></a>Tester le bouton

Enregistrez toutes vos modifications et exécutez `npm start` depuis l’invite de commandes, si le serveur n’est pas déjà en cours d’exécution. Puis procédez comme suit pour tester le bouton **Insérer gist par défaut** bouton.

1. Ouvrez Outlook et rédigez un nouveau message.

1. Dans la fenêtre composer un message, sélectionnez le bouton **Insérer gist par défaut**. Vous devriez voir une boîte de dialogue dans laquelle vous pouvez configurer le complément, en commençant par l’invite de définition de votre nom d’utilisateur GitHub.

    ![Capture d’écran de l’invite de la boîte de dialogue permettant de configurer le complément](../images/addin-prompt-configure.png)

1. Dans la boîte de dialogue Paramètres, entrez votre nom d’utilisateur GitHub, puis soit **Onglet** soit cliquez ailleurs dans la boîte de dialogue pour faire apparaître l’événement `change`, qui devrait charger votre liste de gists publiques. Sélectionnez un gist par défaut, puis cliquez sur **Terminer**.

    ![Capture d’écran de la boîte de dialogue des paramètres du complément](../images/addin-settings.png)

1. Cliquez de nouveau sur le bouton **Insérer un gist par défaut**. Cette fois, le contenu du gist est inséré dans le corps du courrier électronique.

   > [!NOTE]
   > Outlook sur Windows : pour récupérer les paramètres les plus récents, vous devrez peut-être fermer et rouvrir la fenêtre de composition d’un message.

## <a name="implement-a-task-pane"></a>Implémentation d’un volet de tâches

Le bouton de ce complément **Insérer gist** ouvre un volet de tâches et affiche les gists de l’utilisateur. L’utilisateur peut sélectionner un des gists à insérer dans le corps du message. Si l’utilisateur n’a pas encore configuré le complément, il sera invité à le faire.

### <a name="specify-the-html-for-the-task-pane"></a>Spécifier le code HTML pour le volet de tâches

Dans le projet que vous avez créé, le code HTML du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.html**. Ouvrez ce fichier et remplacez l’intégralité de son contenu par le balisage suivant.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Contoso Task Pane Add-in</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l ms-landing-page">
  <main class="ms-landing-page__main">
    <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
      <div id="not-configured" style="display: none;">
        <div class="centered ms-font-xxl ms-u-textAlignCenter">Welcome!</div>
        <div class="ms-font-xl" id="settings-prompt">Please choose the <strong>Settings</strong> icon at the bottom of this window to configure this add-in.</div>
      </div>
      <div id="gist-list-container" style="display: none;">
        <form>
          <div id="gist-list">
          </div>
        </form>
      </div>
      <div id="error-display" style="display: none;" class="ms-u-borderBase ms-fontColor-error ms-font-m ms-bgColor-error ms-borderColor-error">
      </div>
    </section>
    <button class="ms-Button ms-Button--primary" id="insert-button" tabindex=0 disabled>
      <span class="ms-Button-label">Insert</span>
    </button>
  </main>
  <footer class="ms-landing-page__footer ms-bgColor-themePrimary">
    <div class="ms-landing-page__footer--left">
      <img src="../../assets/logo-filled.png" />
      <h1 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-white">Git the gist</h1>
    </div>
    <div id="settings-icon" class="ms-landing-page__footer--right" aria-label="Settings" tabindex=0>
      <i class="ms-Icon enlarge ms-Icon--Settings ms-fontColor-white"></i>
    </div>
  </footer>
  <script type="text/javascript" src="../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../node_modules/showdown/dist/showdown.min.js"></script>
  <script type="text/javascript" src="../node_modules/urijs/src/URI.min.js"></script>
  <script type="text/javascript" src="../src/helpers/addin-config.js"></script>
  <script type="text/javascript" src="../src/helpers/gist-api.js"></script>
  <script type="text/javascript" src="taskpane.js"></script>
</body>

</html>
```

### <a name="specify-the-css-for-the-task-pane"></a>Spécifier le style CSS pour le volet de tâches

Dans le projet que vous avez créé, le style CSS du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.css**. Ouvrez ce fichier et remplacez l’intégralité de son contenu par le code suivant.

```css
/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */
html, body {
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
  overflow: auto; }

body {
  position: relative;
  font-size: 16px; }

main {
  height: 100%;
  overflow-y: auto; }

footer {
  width: 100%;
  position: relative;
  bottom: 0;
  margin-top: 10px;}

p, h1, h2, h3, h4, h5, h6 {
  margin: 0;
  padding: 0; }

ul {
  padding: 0; }

#settings-prompt {
  margin: 10px 0;
}

#error-display {
  padding: 10px;
}

#insert-button {
  margin: 0 10px;
}

.clearfix {
  display: block;
  clear: both;
  height: 0; }

.pointerCursor {
  cursor: pointer; }

.invisible {
  visibility: hidden; }

.undisplayed {
  display: none; }

.ms-Icon.enlarge {
  position: relative;
  font-size: 20px;
  top: 4px; }

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}

.ms-landing-page {
  display: -webkit-flex;
  display: flex;
  -webkit-flex-direction: column;
          flex-direction: column;
  -webkit-flex-wrap: nowrap;
          flex-wrap: nowrap;
  height: 100%; }
  .ms-landing-page__main {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    height: 100%; }

  .ms-landing-page__content {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    height: 100%;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    padding: 20px; }
    .ms-landing-page__content h2 {
      margin-bottom: 20px; }
  .ms-landing-page__footer {
    display: -webkit-inline-flex;
    display: inline-flex;
    -webkit-justify-content: center;
            justify-content: center;
    -webkit-align-items: center;
            align-items: center; }
    .ms-landing-page__footer--left {
      transition: background ease 0.1s, color ease 0.1s;
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: flex-start;
              justify-content: flex-start;
      -webkit-align-items: center;
              align-items: center;
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      padding: 20px; }
      .ms-landing-page__footer--left:active, .ms-landing-page__footer--left:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--left:active {
        background: #005ca4; }
      .ms-landing-page__footer--left--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--left--disabled:active, .ms-landing-page__footer--left--disabled:hover {
          background: transparent; }
      .ms-landing-page__footer--left img {
        width: 40px;
        height: 40px; }
      .ms-landing-page__footer--left h1 {
        -webkit-flex: 1 0 0px;
                flex: 1 0 0px;
        margin-left: 15px;
        text-align: left;
        width: auto;
        max-width: auto;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis; }
    .ms-landing-page__footer--right {
      transition: background ease 0.1s, color ease 0.1s;
      padding: 29px 20px; }
      .ms-landing-page__footer--right:active, .ms-landing-page__footer--right:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--right:active {
        background: #005ca4; }
      .ms-landing-page__footer--right--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--right--disabled:active, .ms-landing-page__footer--right--disabled:hover {
          background: transparent; }
```

### <a name="specify-the-javascript-for-the-task-pane"></a>Spécifier le code JavaScript pour le volet de tâches

Dans le projet que vous avez créé, le code JavaScript du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.js**. Ouvrez ce fichier et remplacez l’intégralité de son contenu par le code suivant.

```js
(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $('#not-configured').show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function(){
        var gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function(gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  {coercionType: Office.CoercionType.Html}, function(result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function(){
        // Display settings dialog.
        var url = new URI('../src/settings/dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        var dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('#insert-button').removeAttr('disabled');
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
```

### <a name="test-the-button"></a>Tester le bouton

Enregistrez toutes vos modifications et exécutez `npm start` depuis l’invite de commandes, si le serveur n’est pas déjà en cours d’exécution. Puis procédez comme suit pour tester le bouton **Insérer gist**.

1. Ouvrez Outlook et rédigez un nouveau message.

1. Dans la fenêtre composer un message, sélectionnez le bouton **Insérer gist**. Vous devriez voir un volet des tâches qui s’ouvre à droite du formulaire Composer.

1. Dans le volet des tâches, sélectionnez le gist **Hello World Html**, puis sélectionnez **insérer** pour insérer ce gist dans le corps du message.

![Capture d’écran du volet Office Complément et du contenu du gist sélectionné qui s’affiche dans le corps du message](../images/addin-taskpane.png)

## <a name="next-steps"></a>Étapes suivantes

Ce didacticiel vous a appris à créer un complément Outlook qui peut être utilisé pour dans le mode composer un message pour insérer du contenu dans le corps d’un message. Pour en savoir plus sur le développement des compléments Outlook, passez à l’article suivant :

> [!div class="nextstepaction"]
> [API de complément Outlook](../outlook/apis.md)
