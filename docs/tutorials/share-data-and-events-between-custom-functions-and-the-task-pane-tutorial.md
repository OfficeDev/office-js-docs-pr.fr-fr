---
title: 'Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office'
description: Découvrez comment partager des données et des événements entre des fonctions personnalisées et le volet Office dans Excel.
ms.date: 06/15/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: b61ac6305586e5de2f53a0950fd6a52a0503eafd
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958727"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office

Partagez des données globales et envoyez des événements entre le volet Office et les fonctions personnalisées de votre complément Excel avec un runtime partagé. Nous vous recommandons d’utiliser un runtime partagé pour la plupart des scénarios de fonctions personnalisées, sauf si vous avez une raison spécifique d’utiliser un complément de fonction uniquement personnalisé. Ce didacticiel suppose que vous êtes familiarisé avec l’utilisation du générateur [Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) pour créer des projets de compléments. Envisagez d’effectuer le [Didacticiel sur les fonctions Excel personnalisées](excel-tutorial-create-custom-functions.md), si ce n’est déjà fait.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Utilisez le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) pour créer le projet de complément Excel.

- Pour générer un complément Excel avec des fonctions personnalisées, exécutez la commande suivante.

    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

Le générateur crée le projet et installe les composants Node de support.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Suivez ces étapes pour configurer le projet de complément pour utiliser un runtime partagé.

1. Démarrez Visual Studio Code et ouvrez le projet de complément que vous avez généré.
1. Ouvrez le fichier **manifest.xml**.
1. Remplacer (ou ajouter) la section suivante XML **\<Requirements\>** pour exiger l'[ensemble d'exigences d'exécution partagé](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    Après la mise à jour, votre manifeste XML doit apparaître dans l'ordre suivant.

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Recherchez la section **\<VersionOverrides\>** et ajoutez la section **\<Runtimes\>** suivante. La durée de vie doit être **longue** afin que votre code de complément puisse s’exécuter même quand le volet Office est fermé. La valeur `resid` est **Taskpane.Url** qui se réfère à l’emplacement du fichier **taskpane.html** spécifiée dans la section `<bt:Urls>` près du bas du fichier **manifest.xml**.

    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```

    > [!IMPORTANT]
    > La section **\<Runtimes\>** doit être entrée après l’élément `<Host xsi:type="...">` dans l’ordre exact indiqué dans le code XML suivant.

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```

    > [!NOTE]
    > Si votre macro complémentaire inclut l’`Runtimes`élément dans le manifeste (runtime partagé requis) et que les conditions d’utilisation de Microsoft Edge avec WebView2 (basées sur Chromium) sont remplies, il utilise ce contrôle WebView2. Si les conditions ne sont pas remplies, il utilise Internet Explorer 11, quelle que soit la version Windows ou Microsoft 365 version. Pour plus d’informations, consultez [Runtimes](/javascript/api/manifest/runtimes) and [Browsers utilisés par les compléments Office ](../concepts/browsers-used-by-office-web-add-ins.md).

1. Recherchez l’élément **\<Page\>** . Puis remplacez l’emplacement de la source **Functions.Page.Url** par **TaskPane.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
      <SourceLocation resid="Taskpane.Url"/>
    </Page>
    ...
    ```

1. Recherchez la balise `<FunctionFile ...>` et remplacez le `resid` de **Commands.Url** par **Taskpane.Url**.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Enregistrez le fichier **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurer le fichier webpack.config.js

Le fichier **webpack.config.js** générera plusieurs chargeurs runtime. Vous devez le modifier pour charger uniquement le runtime JavaScript partagé via le fichier **taskpane.html**.

1. Ouvrez le fichier **webpack.config.js**.
1. Allez dans la `plugins:` rubrique.
1. Supprimez le `functions.html` plugin suivant s'il existe.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Supprimez le `commands.html` plugin suivant s'il existe.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Si vous avez supprimé le `functions` ou `commands` plugin, ajoutez-les en tant que `chunks`. Le code JavaScript suivant affiche l'entrée mise à jour si vous avez supprimé à la fois `functions` et `commands` les plugins.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Enregistrez vos changements et reconstruisez le projet.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Vous pouvez également supprimer les fichiers **functions.html** et **Commands.html**. Le **taskpane.html** charge le **code functions.js** et **Commands.js** dans l'environnement d'exécution JavaScript partagé via les mises à jour du pack Web que vous venez de faire.

1. Enregistrez vos changements et exécutez le projet. Assurez-vous qu'il se charge et s'exécute sans erreur.

   ```command&nbsp;line
   npm run start
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Partager l’état entre une fonction personnalisée et du code du volet Office

À présent que les fonctions personnalisées s’exécutent dans le même contexte que votre code du volet Office, elles peuvent partager l’état directement, sans utiliser l’objet **Storage**. Les instructions suivantes montrent comment partager une variable globale entre une fonction personnalisée et du code du volet Office.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Créer des fonctions personnalisées pour obtenir ou stocker l’état partagé

1. Dans Visual Studio Code, ouvrez le fichier **src/functions/functions.js**.
1. Sur la ligne 1, tout en haut, insérez le code suivant. Cette opération initialise une variable globale nommée **sharedState**.

    ```js
    window.sharedState = "empty";
    ```

1. Ajoutez le code suivant pour créer une fonction personnalisée qui stocke des valeurs dans la variable **sharedState**.

    ```js
    /**
     * Saves a string value to shared state with the task pane
     * @customfunction STOREVALUE
     * @param {string} value String to write to shared state with task pane.
     * @return {string} A success value
     */
    function storeValue(sharedValue) {
      window.sharedState = sharedValue;
      return "value stored";
    }
    ```

1. Ajoutez le code suivant pour créer une fonction personnalisée qui obtient la valeur actuelle de la variable **sharedState**.

    ```js
    /**
     * Gets a string value from shared state with the task pane
     * @customfunction GETVALUE
     * @returns {string} String value of the shared state with task pane.
     */
    function getValue() {
      return window.sharedState;
    }
    ```

1. Enregistrez le fichier.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Créer des contrôles du volet Office pour utiliser des données globales

1. Ouvrez le fichier **src/taskpane/taskpane.html**.
1. Ajoutez l’élément de script suivant juste avant l’élément de fermeture `</head>`.

    ```HTML
    <script src="../functions/functions.js"></script>
    ```

1. Après l’élément de fermeture `</main>`, ajoutez le code HTML suivant. Le code HTML crée deux zones de texte et des boutons permettant d’obtenir ou de stocker des données globales.

    ```HTML
    <ol>
      <li>
        Enter a value to send to the custom function and select
        <strong>Store</strong>.
      </li>
      <li>
        Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
      </li>
      <li>
        To send data to the task pane, in a cell, enter
        <strong>=CONTOSO.STOREVALUE("new value")</strong>
      </li>
      <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>

    <p>Store new value to shared state</p>
    <div>
      <input type="text" id="storeBox" />
      <button onclick="storeSharedValue()">Store</button>
    </div>

    <p>Get shared state value</p>
    <div>
      <input type="text" id="getBox" />
      <button onclick="getSharedValue()">Get</button>
    </div>
    ```

1. Avant `</body>`l'élément de fermeture, ajoutez le script suivant. Ce code gérera les événements de clic sur le bouton lorsque l'utilisateur souhaite stocker ou obtenir des données globales.

    ```HTML
    <script>
      function storeSharedValue() {
        let sharedValue = document.getElementById('storeBox').value;
        window.sharedState = sharedValue;
      }

      function getSharedValue() {
        document.getElementById('getBox').value = window.sharedState;
      }
   </script>
   ```

1. Enregistrez le fichier.
1. Créez le projet.

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Essayer de partager des données entre les fonctions personnalisées et le volet Office

- Démarrez le projet à l’aide de la commande suivante.

    ```command line
    npm run start
    ```

Une fois Excel démarré, vous pouvez utiliser les boutons du volet Office pour stocker ou obtenir des données partagées. Entrez `=CONTOSO.GETVALUE()` dans une cellule pour que la fonction personnalisée extraie les mêmes données partagées. Vous pouvez également utiliser `=CONTOSO.STOREVALUE("new value")` pour remplacer les données partagées par une nouvelle valeur.

> [!NOTE]
> La configuration de votre projet comme illustré dans cet article permet de partager le contexte entre des fonctions personnalisées et le volet Office. L’appel de certaines API Office à partir de fonctions personnalisées est possible. Pour plus d’informations, [consultez Appeler les API Microsoft Excel à partir d’une fonction personnalisée](../excel/call-excel-apis-from-custom-function.md).

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
